import chokidar from "chokidar";

export type WatchTrigger = {
  kind: "initial" | "change";
  paths: readonly string[];
};

type WatchRunEnvelope<T> =
  | T
  | { value: T; watchPaths?: readonly string[] };

export type WatchControllerOptions<T> = {
  watchPaths: readonly string[];
  debounceMs: number;
  watchOptions?: Parameters<typeof chokidar.watch>[1];
  runInitial?: boolean;
  onReady?: () => Promise<void> | void;
  run(trigger: WatchTrigger): Promise<WatchRunEnvelope<T>>;
  emitResult(result: T, trigger: WatchTrigger): Promise<void> | void;
  emitError(error: unknown, trigger: WatchTrigger): Promise<void> | void;
  signal(
    signal: "SIGINT" | "SIGTERM",
    listener: () => void,
  ): void;
  onStop?(reason: "SIGINT" | "SIGTERM"): Promise<void> | void;
};

const isRunEnvelope = <T>(
  value: WatchRunEnvelope<T>,
): value is { value: T; watchPaths?: readonly string[] } =>
  typeof value === "object" &&
  value !== null &&
  Array.isArray(value) === false &&
  "value" in value;

/** Runs the shared chokidar/debounce/ordering/signal lifecycle for a watch transport. */
export const runWatchController = async <T>({
  watchPaths: initialPaths,
  debounceMs,
  watchOptions,
  runInitial = true,
  onReady,
  run,
  emitResult,
  emitError,
  signal,
  onStop,
}: WatchControllerOptions<T>): Promise<number> => {
  let paths = [...initialPaths];
  const watcher = chokidar.watch(paths, {
    ignoreInitial: true,
    ...watchOptions,
  });
  await new Promise<void>((resolve, reject) => {
    watcher.once("ready", resolve).once("error", reject);
  });
  let closing = false;
  let timer: NodeJS.Timeout | undefined;
  let running = false;
  let dirtyTrigger: WatchTrigger | undefined;
  let inFlight: Promise<void> | undefined;
  const applyWatchedPaths = async (next: readonly string[]): Promise<void> => {
    const nextPaths = [...new Set(next)];
    const removed = paths.filter((path) => !nextPaths.includes(path));
    const added = nextPaths.filter((path) => !paths.includes(path));
    if (removed.length) await watcher.unwatch(removed);
    if (added.length) watcher.add(added);
    paths = nextPaths;
  };
  const execute = async (trigger: WatchTrigger): Promise<void> => {
    if (closing) return;
    if (running) {
      dirtyTrigger = trigger;
      return;
    }
    running = true;
    try {
      const envelope = await run(trigger);
      const result = isRunEnvelope(envelope) ? envelope.value : envelope;
      if (isRunEnvelope(envelope) && envelope.watchPaths !== undefined)
        await applyWatchedPaths(envelope.watchPaths);
      if (!closing) await emitResult(result, trigger);
    } catch (error) {
      if (!closing) await emitError(error, trigger);
    } finally {
      running = false;
      if (dirtyTrigger !== undefined && !closing) {
        const next = dirtyTrigger;
        dirtyTrigger = undefined;
        start(next);
      }
    }
  };
  const start = (trigger: WatchTrigger): void => {
    const task = execute(trigger);
    inFlight = task;
    void task;
  };
  const queue = (trigger: WatchTrigger): void => {
    if (closing) return;
    clearTimeout(timer);
    timer = setTimeout(() => {
      timer = undefined;
      start(trigger);
    }, debounceMs);
  };
  if (onReady) await onReady();
  if (runInitial) {
    start({ kind: "initial", paths: [...paths] });
    await inFlight;
  }
  watcher.on("all", (_event, changed) =>
    queue({ kind: "change", paths: [String(changed)] }),
  );
  watcher.on("error", (error) => {
    if (!closing)
      void emitError(error, { kind: "change", paths: [] });
  });
  const completion = Promise.withResolvers<number>();
  const stop = async (reason: "SIGINT" | "SIGTERM"): Promise<void> => {
    if (closing) return;
    closing = true;
    dirtyTrigger = undefined;
    clearTimeout(timer);
    try {
      if (inFlight) await inFlight;
      await watcher.close();
      if (onStop) await onStop(reason);
      completion.resolve(reason === "SIGINT" ? 130 : 143);
    } catch (error) {
      completion.reject(error);
    }
  };
  signal("SIGINT", () => void stop("SIGINT"));
  signal("SIGTERM", () => void stop("SIGTERM"));
  return completion.promise;
};
