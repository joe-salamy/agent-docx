export type ProjectContext = {
  manifestPath: string;
  clock: () => Date;
  randomUuid: () => string;
};
