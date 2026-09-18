import {
  createContext,
  useContext,
  useSyncExternalStore,
  type ReactNode,
} from "react";
import type { Workbook, WorkbookState } from "@nori-internal/core";
const ReadOnlyContext = createContext(false);
/** Whether this view permits document edits; runtime APIs remain available to the host. */
export function useWorkbookReadOnly(): boolean {
  return useContext(ReadOnlyContext);
}
const Context = createContext<Workbook | null>(null);
/** Provide a live workbook to composable DOM primitives. */
export function WorkbookProvider({
  workbook,
  children,
  readOnly = false,
}: {
  readonly workbook: Workbook;
  readonly readOnly?: boolean;
  readonly children: ReactNode;
}): ReactNode {
  return (
    <ReadOnlyContext.Provider value={readOnly}>
      <Context.Provider value={workbook}>{children}</Context.Provider>
    </ReadOnlyContext.Provider>
  );
}
/** Access the workbook command API inside a provider. */
export function useWorkbook(): Workbook {
  const workbook = useContext(Context);
  if (!workbook) throw new Error("Nori components require WorkbookProvider");
  return workbook;
}
/** React external-store subscription, safe for SSR with a stable runtime instance. */
export function useWorkbookState(): WorkbookState {
  const workbook = useWorkbook();
  return useSyncExternalStore(
    workbook.subscribe,
    workbook.getState,
    workbook.getState,
  );
}
