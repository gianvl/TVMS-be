import { Document } from 'mongoose';
import { z } from 'zod';

export interface IDriver {
  lastName: string | null;
  firstName: string | null;
}

export interface IConfiscatedItem {
  type: string | null;
  number: string | null;
}

export interface IApprehension {
  dateOfSubmission: Date | null;
  daysInterval: number | null;
  dateOfApprehension: Date | null;
  timeOfApprehension: string | null;
  agency: string | null;
  apprehendingOfficer: string | null;
  caseNumber: string | null;
  driver: IDriver;
  violation: string | null;
  confiscatedItem: IConfiscatedItem;
  restrictionCode: string | null;
  conditions: string | null;
  nationality: string | null;
  gender: string | null;
  mvType: string | null;
  plateNumber: string | null;
  placeOfApprehension: string | null;
  remarks: string | null;
}

export interface IApprehensionDocument extends IApprehension, Document {}

// Zod schemas for validation
const driverSchema = z.object({
  lastName: z.string().min(1, 'Driver last name is required'),
  firstName: z.string().min(1, 'Driver first name is required'),
}).strict();

const confiscatedItemSchema = z.object({
  type: z.string().nullable().optional(),
  number: z.string().nullable().optional(),
}).strict();

export const createApprehensionSchema = z.object({
  dateOfSubmission: z.coerce.date(),
  dateOfApprehension: z.coerce.date(),
  agency: z.string().min(1, 'Agency is required'),
  apprehendingOfficer: z.string().min(1, 'Apprehending officer is required'),
  caseNumber: z.string().min(1, 'Case number is required'),
  driver: driverSchema,
  violation: z.string().min(1, 'Violation is required'),
  plateNumber: z.string().min(1, 'Plate number is required'),
  placeOfApprehension: z.string().min(1, 'Place of apprehension is required'),
  // Optional fields
  timeOfApprehension: z.string().nullable().optional(),
  mvType: z.string().nullable().optional(),
  confiscatedItem: confiscatedItemSchema.optional(),
  restrictionCode: z.string().nullable().optional(),
  conditions: z.string().nullable().optional(),
  nationality: z.string().nullable().optional(),
  gender: z.string().nullable().optional(),
  remarks: z.string().nullable().optional(),
}).strict();

export const updateApprehensionSchema = z.object({
  dateOfSubmission: z.coerce.date().optional(),
  dateOfApprehension: z.coerce.date().optional(),
  timeOfApprehension: z.string().min(1).optional(),
  agency: z.string().min(1).optional(),
  apprehendingOfficer: z.string().min(1).optional(),
  caseNumber: z.string().min(1).optional(),
  driver: driverSchema.optional(),
  violation: z.string().min(1).optional(),
  confiscatedItem: confiscatedItemSchema.optional(),
  mvType: z.string().min(1).optional(),
  plateNumber: z.string().min(1).optional(),
  placeOfApprehension: z.string().min(1).optional(),
  restrictionCode: z.string().nullable().optional(),
  conditions: z.string().nullable().optional(),
  nationality: z.string().nullable().optional(),
  gender: z.string().nullable().optional(),
  remarks: z.string().nullable().optional(),
}).strict().refine(data => Object.keys(data).length > 0, {
  message: 'At least one field must be provided for update',
});

export type CreateApprehensionInput = z.infer<typeof createApprehensionSchema>;
export type UpdateApprehensionInput = z.infer<typeof updateApprehensionSchema>;

export interface ImportResult {
  total: number;
  imported: number;
  skipped: number;
}

export interface BulkImportRowError {
  row: number;
  error: string;
}

export interface BulkImportResult {
  total: number;
  imported: number;
  failed: number;
  errors: BulkImportRowError[];
}

export interface ApprehensionFilters {
  dateFrom?: Date;
  dateTo?: Date;
  agency?: string;
  violation?: string;
  mvType?: string;
  plateNumber?: string;
  driverName?: string;
}

export interface StatsFilters {
  month?: string;           // YYYY-MM format
  dateFrom?: Date;
  dateTo?: Date;
  agency?: string;
  violation?: string;
  placeOfApprehension?: string;
  topLimit?: number;
}

export interface StatsResponse {
  total: number;
  topAgencies: { agency: string; count: number }[];
  topViolations: { violation: string; count: number }[];
  topLocations: { location: string; count: number }[];
}
