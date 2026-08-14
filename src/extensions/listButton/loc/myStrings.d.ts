declare interface IListButtonCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ListButtonCommandSetStrings' {
  const strings: IListButtonCommandSetStrings;
  export = strings;
}
