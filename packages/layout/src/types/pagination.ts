import type { Box } from "@pptkit/core";

export interface TablePaginationRule {
  /** A top-level table on a source slide. At most one paginated table per slide. */
  elementId: string;
  /** Available slide-local rectangle for the first fragment. Width does not resize columns. */
  area: Readonly<Box>;
  headerRows?: number;
  continuation?: { layoutId?: string; area?: Readonly<Box> };
  /** preserve (default): reject the rule; overflow: place an oversized block alone with a warning. */
  overflow?: "preserve" | "overflow";
  /** Preserve native first/last-row boundaries across fragments; native bands can still restart. */
  nativeRowRegions?: "fragment" | "source-boundaries";
  includeHidden?: boolean;
  /** Includes the first fragment. Default 1000. */
  maxPages?: number;
}

export interface TablePaginationFragment {
  sourceSlideId: string;
  sourceElementId: string;
  slideId: string;
  elementId: string;
  /** Zero-based fragment ordinal; zero retains source slide/table identity. */
  pageIndex: number;
  sourceRowIndices: number[];
  /** Original header rows on fragment zero are not repetitions. */
  repeatedHeaderRowCount: number;
}

export interface LayoutPagination {
  fragments: TablePaginationFragment[];
}
