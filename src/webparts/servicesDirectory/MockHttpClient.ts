import { ServiceDirectory,ServiceDirectorys } from './ServiceDirectoryList';


export default class MockHttpClient  {

   private static _items: ServiceDirectory[] = [
       { Title: 'Mock List', ID: 1 ,AverageRating:0,Contact:"",Description:"",LocationMap:null,Logo:"",Phone:"",ServiceType:"",Website:""},
       { Title: 'Mock List2', ID: 2 ,AverageRating:2.5,Contact:"",Description:"",LocationMap:null,Logo:"",Phone:"",ServiceType:"",Website:""},
       { Title: 'Mock List3', ID: 3 ,AverageRating:0,Contact:"",Description:"",LocationMap:null,Logo:"",Phone:"",ServiceType:"",Website:""}
    ];

   public static get(): Promise<ServiceDirectory[]> {
   return new Promise<ServiceDirectory[]>((resolve) => {
           resolve(MockHttpClient._items);
       });
   }
}