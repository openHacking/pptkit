import type { OpcInspectionOptions } from "./opc.js";
import type { PptxSourceLocation } from "./source.js";
export interface MasterPictureInventoryLimits { readonly maxWorkUnits: number; readonly maxRecords: number; readonly maxTextCharacters: number }
export interface MasterPictureInventoryOptions extends OpcInspectionOptions { readonly masterPictureLimits?: Partial<MasterPictureInventoryLimits> }
export interface MasterPictureRelationship {
  readonly id: string;
  readonly status: "internal-image" | "external-image" | "missing" | "wrong-type" | "fragmented";
  readonly targetPart?: string;
}
export interface MasterPictureInventoryItem {
  readonly source: PptxSourceLocation;
  readonly sourceObjectId?: string;
  readonly name?: string;
  /** Raw XML indices of containing groups, ordered from root to immediate parent. */
  readonly groupElementIndices: readonly number[];
  readonly placeholder?: { readonly type?: string; readonly index?: string };
  readonly embed?: MasterPictureRelationship;
  readonly link?: MasterPictureRelationship;
  /** Original lexical values only; no interpretation or inherited geometry. */
  readonly transform?: { readonly x?: string; readonly y?: string; readonly cx?: string; readonly cy?: string; readonly rotation?: string; readonly flipH?: string; readonly flipV?: string };
  readonly crop?: { readonly left?: string; readonly top?: string; readonly right?: string; readonly bottom?: string };
}
export interface MasterPicturePartInventory {
  readonly part: string;
  readonly sourceMasterId: number;
  readonly layoutParts: readonly string[];
  /** Source slide IDs referencing this master through a listed layout, not visibility claims. */
  readonly slideIds: readonly number[];
  /** Raw inheritance switches for listed-layout paths; absence is not interpreted as a default. */
  readonly slideUses: readonly { readonly slideId: number; readonly slidePart: string; readonly layoutPart: string; readonly slideShowMasterSp?: string; readonly layoutShowMasterSp?: string }[];
  readonly pictures: readonly MasterPictureInventoryItem[];
  readonly otherDrawingCount: number;
}
export interface MasterPictureInventory {
  readonly schemaVersion: "pptkit-master-picture-inventory-v1";
  readonly dialect: "strict" | "transitional";
  readonly masters: readonly MasterPicturePartInventory[];
  readonly workUnits: number;
  readonly recordCount: number;
  readonly textCharacters: number;
  readonly coverage: "effective-listed-masters-only";
  readonly mapping: "not-performed";
}
