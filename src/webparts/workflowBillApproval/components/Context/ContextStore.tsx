import { WebPartContext } from "@microsoft/sp-webpart-base";
import { createContext } from "react";

interface IContextStore {
  spContext: WebPartContext;
}
export const ContextStore = createContext<IContextStore>({
  spContext: {} as WebPartContext,
});
