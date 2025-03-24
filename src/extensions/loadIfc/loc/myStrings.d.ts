declare interface ILoadIfcCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'LoadIfcCommandSetStrings' {
  const strings: ILoadIfcCommandSetStrings;
  export = strings;
}
