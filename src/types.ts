export interface ExtractedData {
  headers: string[];
  rows: string[][];
  usage?: {
    promptTokens: number;
    responseTokens: number;
  };
}

export interface AppState {
  step: "UPLOAD" | "PROCESSING" | "RESULTS";
  data: ExtractedData | null;
  error: string | null;
}
