import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from '@microsoft/sp-http';
import { CompactPeoplePicker, IPersonaProps, IBasePickerSuggestionsProps } from 'office-ui-fabric-react';
import { PageContext } from '@microsoft/sp-page-context'; 
export interface IOrganisationalChartProps {
  description: string;
  listName: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  defaultValues?: IPersonaProps[];
  multi?: boolean;
  onChange?(people: IPersonaProps[]): void;
  numberOfItems:number;
  pageContext: PageContext;// here we passing page context


}
