import { AppSettings, DEFAULT_SETTINGS } from "../models/settings";

const STORAGE_KEY = "slidesci_settings";

export class SettingsService {
  private settings: AppSettings;

  constructor() {
    this.settings = this.load();
  }

  private load(): AppSettings {
    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      if (stored) {
        const parsed = JSON.parse(stored) as Partial<AppSettings>;
        return this.mergeWithDefaults(parsed);
      }
    } catch {
      // ignore parse errors
    }
    return { ...DEFAULT_SETTINGS };
  }

  private mergeWithDefaults(partial: Partial<AppSettings>): AppSettings {
    return {
      title: { ...DEFAULT_SETTINGS.title, ...partial.title },
      label: { ...DEFAULT_SETTINGS.label, ...partial.label },
      arrange: { ...DEFAULT_SETTINGS.arrange, ...partial.arrange },
      codeBlock: { ...DEFAULT_SETTINGS.codeBlock, ...partial.codeBlock },
      export: { ...DEFAULT_SETTINGS.export, ...partial.export },
    };
  }

  save(): void {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(this.settings));
    } catch {
      // storage full or unavailable
    }
  }

  getSettings(): AppSettings {
    return this.settings;
  }

  updateTitle(updates: Partial<AppSettings["title"]>): void {
    this.settings.title = { ...this.settings.title, ...updates };
    this.save();
  }

  updateLabel(updates: Partial<AppSettings["label"]>): void {
    this.settings.label = { ...this.settings.label, ...updates };
    this.save();
  }

  updateArrange(updates: Partial<AppSettings["arrange"]>): void {
    this.settings.arrange = { ...this.settings.arrange, ...updates };
    this.save();
  }

  updateCodeBlock(updates: Partial<AppSettings["codeBlock"]>): void {
    this.settings.codeBlock = { ...this.settings.codeBlock, ...updates };
    this.save();
  }

  updateExport(updates: Partial<AppSettings["export"]>): void {
    this.settings.export = { ...this.settings.export, ...updates };
    this.save();
  }

  resetToDefaults(): void {
    this.settings = { ...DEFAULT_SETTINGS };
    this.save();
  }
}

export const settingsService = new SettingsService();
