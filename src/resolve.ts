import { liberationSerifBoldBase64, liberationSerifBoldItalicBase64, liberationSerifItalicBase64, liberationSerifRegularBase64 } from "./bundled-fonts.js";
import { createHash } from "node:crypto";
import * as fontkit from "fontkit";
import type { Font } from "fontkit";
import { builtInProfiles } from "./profiles.js";
import { AgentDocxError, type Diagnostic, type EstimateOptions, type FontSetInput, type LayoutProfile, type MetricFont, type ResolvedLayoutProfile, type TextStyle } from "./types.js";
export type LoadedFace={bytes:Uint8Array;hash:string;font:Font};
export type LoadedFonts={regular:LoadedFace;bold:LoadedFace;italic:LoadedFace;boldItalic:LoadedFace;family:string;metrics:MetricFont[];warnings:Diagnostic[]};
const roles=["regular","bold","italic","boldItalic"] as const;
const sha=(b:Uint8Array)=>createHash("sha256").update(b).digest("hex");
const clone=<T>(v:T):T=>structuredClone(v);
const mergeStyle=(base:TextStyle,over?:Partial<TextStyle>):TextStyle=>({...base,...over,lineSpacing:over?.lineSpacing??base.lineSpacing});
export function resolveProfile(options:EstimateOptions):LayoutProfile{
  let selected:LayoutProfile;
  if(typeof options.profile==="string"){
    selected=builtInProfiles[options.profile];
    if(!selected)throw new AgentDocxError("INVALID_ARGUMENT",`Unknown profile: ${options.profile}`);
  }else selected=options.profile??builtInProfiles["us-district-conventional"];
  const p=clone(selected);
  const original=JSON.stringify({page:p.page,requestedFontFamily:p.requestedFontFamily,body:p.body,headings:p.headings,blockquote:p.blockquote,list:p.list,footnote:p.footnote});
  const apply=(o:typeof options.layout|undefined,source:"template"|"override")=>{
    if(!o)return;
    if(o.page)p.page={...p.page,...o.page,marginsTwips:{...p.page.marginsTwips,...o.page.marginsTwips}};
    if(o.requestedFontFamily!==undefined)p.requestedFontFamily=o.requestedFontFamily;
    if(o.body)p.body=mergeStyle(p.body,o.body);
    if(o.headings)for(const key of ["1","2","3","4","5","6"] as const)if(o.headings[key])p.headings={...p.headings,[key]:mergeStyle(p.headings[key],o.headings[key])};
    if(o.blockquote)p.blockquote=mergeStyle(p.blockquote,o.blockquote);
    if(o.list)p.list=mergeStyle(p.list,o.list);
    if(o.footnote)p.footnote=mergeStyle(p.footnote,o.footnote);
    if(o.pagination)p.pagination={...p.pagination,...o.pagination};
    p.provenance={...p.provenance,"":{source}};
  };
  apply(options.template?.imported,"template");
  apply(options.layout,"override");
  const actual=JSON.stringify({page:p.page,requestedFontFamily:p.requestedFontFamily,body:p.body,headings:p.headings,blockquote:p.blockquote,list:p.list,footnote:p.footnote});
  if((p.id==="frap-32"||p.id==="cand-civil")&&actual!==original)p.warnings=[...p.warnings,{code:"PROFILE_CONSTRAINT_VIOLATION",severity:"warning",message:"Template or explicit layout values differ from the selected court profile; actual values are preserved and filing compliance must be verified."}];
  validateProfile(p);
  return p;
}
function finite(name:string,n:number,min:number,integer=false){if(!Number.isFinite(n)||n<min||(integer&&!Number.isInteger(n)))throw new AgentDocxError("INVALID_LAYOUT",`${name} must be ${integer?"an integer ":""}>= ${min}`)}
export function validateProfile(p:LayoutProfile){finite("page.widthTwips",p.page.widthTwips,1);finite("page.heightTwips",p.page.heightTwips,1);for(const [k,v] of Object.entries(p.page.marginsTwips))finite(`page.marginsTwips.${k}`,v,0);finite("headerTwips",p.page.headerTwips,0);finite("footerTwips",p.page.footerTwips,0);finite("gutterTwips",p.page.gutterTwips,0);if(p.page.widthTwips-p.page.marginsTwips.left-p.page.marginsTwips.right-p.page.gutterTwips<=0||p.page.heightTwips-p.page.marginsTwips.top-p.page.marginsTwips.bottom<=0)throw new AgentDocxError("INVALID_LAYOUT","Resolved page has non-positive usable area");for(const s of [p.body,...Object.values(p.headings),p.blockquote,p.list,p.footnote]){finite("fontSizeTwips",s.fontSizeTwips,1);for(const n of [s.beforeTwips,s.afterTwips,s.leftIndentTwips,s.rightIndentTwips,s.firstLineIndentTwips,s.hangingIndentTwips])finite("style geometry",n,0);if(s.lineSpacing.rule==="auto"){finite("lineSpacing.numerator",s.lineSpacing.numerator,1);if(s.lineSpacing.denominator!==240)throw new AgentDocxError("INVALID_LAYOUT","Auto line spacing denominator must be 240")}else finite("lineSpacing.twips",s.lineSpacing.twips,1)}finite("widowLines",p.pagination.widowLines,1,true);finite("orphanLines",p.pagination.orphanLines,1,true);if(p.pagination.maxCountedLinesPerPage!==null)finite("maxCountedLinesPerPage",p.pagination.maxCountedLinesPerPage,1,true)}
async function bundled():Promise<FontSetInput>{return{family:"Liberation Serif",regular:Buffer.from(liberationSerifRegularBase64,"base64"),bold:Buffer.from(liberationSerifBoldBase64,"base64"),italic:Buffer.from(liberationSerifItalicBase64,"base64"),boldItalic:Buffer.from(liberationSerifBoldItalicBase64,"base64")}}
export async function loadFonts(input:FontSetInput|undefined,requested:string):Promise<LoadedFonts>{
  const source=input??await bundled();
  const warnings:Diagnostic[]=[];
  const createdRegular=fontkit.create(Buffer.from(source.regular));
  if("fonts" in createdRegular)throw new AgentDocxError("INVALID_FONT","Font collections are not accepted; supply a single font face");
  const regularFont:Font=createdRegular;
  const embedded=regularFont.familyName??source.family;
  if(input&&embedded.toLocaleLowerCase("en-US")!==source.family.toLocaleLowerCase("en-US"))throw new AgentDocxError("INVALID_FONT",`Regular font family ${embedded} does not match ${source.family}`);
  const faces={} as Record<(typeof roles)[number],LoadedFace>;
  for(const role of roles){
    const supplied=source[role];
    const bytes=supplied??source.regular;
    if(!supplied&&role!=="regular")warnings.push({code:"FONT_STYLE_FACE_REUSED",severity:"warning",message:`The regular face is reused for omitted ${role} metrics.`,details:{role}});
    const created=role==="regular"?regularFont:fontkit.create(Buffer.from(bytes));
    if("fonts" in created)throw new AgentDocxError("INVALID_FONT",`${role} is a font collection; supply a single font face`);
    const font:Font=created;
    const family=font.familyName??embedded;
    if(input&&family.toLocaleLowerCase("en-US")!==embedded.toLocaleLowerCase("en-US"))throw new AgentDocxError("INVALID_FONT",`${role} face family ${family} differs from regular face ${embedded}`);
    faces[role]={bytes,hash:sha(bytes),font};
  }
  const metrics=roles.map(role=>({role,requestedFamily:requested,metricsFamily:input?source.family:"Liberation Serif",sha256:faces[role].hash,substitutedMetrics:!input&&requested.toLocaleLowerCase("en-US")!=="liberation serif"}));
  return{...faces,family:source.family,metrics,warnings};
}
export function resolvedProfile(profile:LayoutProfile,fonts:LoadedFonts,options:EstimateOptions):ResolvedLayoutProfile{return{...profile,metricFonts:fonts.metrics,template:options.template?{packageSha256:options.template.package.sha256,mainPart:options.template.package.mainPart,selectedSection:options.template.selectedSection,macroEnabled:options.template.package.macroEnabled,warnings:options.template.warnings}:null}}
