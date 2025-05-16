import { createContext } from "react";

interface IContextStore {
  description: string;
}
export const ContextStore = createContext<IContextStore>({ description: "" });
