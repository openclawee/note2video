/// <reference types="vite/client" />

export type N2vApi = {
  openProjectDir: () => Promise<string | null>;
  readText: (filePath: string) => Promise<string>;
  writeText: (filePath: string, text: string) => Promise<void>;
  openPath: (targetPath: string) => Promise<string | null>;
  joinPath: (root: string, ...parts: string[]) => Promise<string>;
  toFileUrl: (filePath: string) => Promise<string>;
};

declare global {
  interface Window {
    n2v?: N2vApi;
  }
}
