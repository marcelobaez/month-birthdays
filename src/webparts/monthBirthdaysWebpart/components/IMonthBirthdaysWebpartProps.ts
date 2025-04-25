import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ITestWebpartNode22Props {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  maxDisplayNumber: number;
  isEditMode: boolean;
  title: string;
  setTitle: (newTitle: string) => void;
}
