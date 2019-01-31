import
{
  ISPNewsItem
}
from './NewsFeedWebPart';

export default class MockHttpClient {  
  private static _items: ISPNewsItem[] = [{ ID:'', TitleEnglish: '', TitleFrench: '', ContentEnglish: '',ContentFrench: '' },];
  public static get(restUrl: string, options?: any): Promise<ISPNewsItem[]>
  {
    return new Promise<ISPNewsItem[]>((resolve) =>
    {
      resolve(MockHttpClient._items);
    });
  }
}