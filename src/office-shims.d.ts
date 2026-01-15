declare namespace Office {
  enum AsyncResultStatus {
    Succeeded,
    Failed
  }

  enum HostType {
    Word,
    Excel,
    PowerPoint
  }

  enum EventType {
    DocumentSelectionChanged
  }

  interface AsyncResult<T> {
    status: AsyncResultStatus;
    value: T;
    error?: { message?: string };
  }

  interface RoamingSettings {
    get(key: string): unknown;
    set(key: string, value: unknown): void;
    saveAsync(callback?: (result: AsyncResult<void>) => void): void;
  }

  interface Document {
    addHandlerAsync(
      eventType: EventType,
      handler: (args: unknown) => void,
      callback?: (result: AsyncResult<void>) => void
    ): void;
  }

  interface Context {
    host: HostType;
    roamingSettings: RoamingSettings;
    document: Document;
  }

  const context: Context;

  function onReady(callback: () => void | Promise<void>): void;
}

declare namespace Word {
  interface RequestContext {
    document: any;
    sync(): Promise<void>;
  }

  function run<T>(callback: (context: RequestContext) => Promise<T>): Promise<T>;
}

declare namespace PowerPoint {
  interface RequestContext {
    presentation: any;
    sync(): Promise<void>;
  }

  function run<T>(callback: (context: RequestContext) => Promise<T>): Promise<T>;
}

declare namespace Excel {
  interface RequestContext {
    workbook: any;
    sync(): Promise<void>;
  }

  function run<T>(callback: (context: RequestContext) => Promise<T>): Promise<T>;
}
