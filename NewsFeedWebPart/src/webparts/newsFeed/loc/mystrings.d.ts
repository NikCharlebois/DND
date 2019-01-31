declare interface INewsFeedWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  NewsFeedNameFieldLabel: string;
  TitleEnglishColumn: string;
  TitleFrenchColumn: string;
  ContentEnglishColumn: string;
  ContentFrenchColumn: string;
}

declare module 'NewsFeedWebPartStrings' {
  const strings: INewsFeedWebPartStrings;
  export = strings;
}
