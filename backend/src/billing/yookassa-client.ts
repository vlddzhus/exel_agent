import { randomUUID } from "crypto";

export interface YooKassaPayment {
  id: string;
  status: string;
  confirmation?: {
    confirmation_url?: string;
  };
  payment_method?: {
    id?: string;
    saved?: boolean;
  };
}

export class YooKassaApiError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly body: string,
  ) {
    super(message);
  }
}

export class YooKassaClient {
  private readonly baseUrl = "https://api.yookassa.ru/v3";
  private readonly authHeader: string;

  constructor(shopId: string, secretKey: string) {
    this.authHeader = `Basic ${Buffer.from(`${shopId}:${secretKey}`).toString(
      "base64",
    )}`;
  }

  createPayment(
    body: Record<string, unknown>,
    idempotenceKey: string = randomUUID(),
  ): Promise<YooKassaPayment> {
    return this.request<YooKassaPayment>("/payments", {
      method: "POST",
      body,
      idempotenceKey,
    });
  }

  getPayment(paymentId: string): Promise<YooKassaPayment> {
    return this.request<YooKassaPayment>(
      `/payments/${encodeURIComponent(paymentId)}`,
      { method: "GET" },
    );
  }

  deactivateSavedPayment(
    paymentMethodId: string,
    idempotenceKey: string = randomUUID(),
  ): Promise<unknown> {
    return this.request(
      `/saved_payment_methods/${encodeURIComponent(paymentMethodId)}/deactivate`,
      { method: "POST", body: {}, idempotenceKey },
    );
  }

  private async request<T>(
    path: string,
    options: {
      method: "GET" | "POST";
      body?: Record<string, unknown>;
      idempotenceKey?: string;
    },
  ): Promise<T> {
    const response = await fetch(`${this.baseUrl}${path}`, {
      method: options.method,
      headers: {
        Authorization: this.authHeader,
        "Content-Type": "application/json",
        ...(options.idempotenceKey
          ? { "Idempotence-Key": options.idempotenceKey }
          : {}),
      },
      body: options.body ? JSON.stringify(options.body) : undefined,
    });

    const text = await response.text();
    if (!response.ok) {
      throw new YooKassaApiError(
        `YooKassa request failed with status ${response.status}`,
        response.status,
        text,
      );
    }

    return (text ? JSON.parse(text) : {}) as T;
  }
}
