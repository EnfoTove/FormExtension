declare interface IHelloWorldFormCustomizerStrings {
  Save: string;
  Cancel: string;
  Close: string;
  Title: string;
}

declare module 'HelloWorldFormCustomizerStrings' {
  const strings: IHelloWorldFormCustomizerStrings;
  export = strings;
}
