export interface ServiceDirectorys {
  value: ServiceDirectory[];
 }
 
 export interface ServiceDirectory {
  Title: string;
  ID: number;
  Description:string;
  LocationMap: string;
  ServiceType: string;
  Website: string;
  AverageRating: number;
  Phone: string;
  Logo: string;
  Contact: string;
 }