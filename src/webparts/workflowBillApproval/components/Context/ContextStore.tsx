import { WebPartContext } from "@microsoft/sp-webpart-base";
import { createContext } from "react";

interface IContextStore {
  context: WebPartContext;
}
export const ContextStore = createContext<IContextStore>({
  context: {} as WebPartContext,
});
