/**
 * BridgeHttpAdapter — a virtual IHttpServerAdapter that captures the route
 * handler registered by App.initialize() and exposes dispatch() for
 * handleWebhook() to call.  We never own the HTTP server.
 */

import type {
  HttpMethod,
  HttpRouteHandler,
  IHttpServerAdapter,
  IHttpServerRequest,
  IHttpServerResponse,
} from "@microsoft/teams.apps";

export class BridgeHttpAdapter implements IHttpServerAdapter {
  private handler: HttpRouteHandler | null = null;

  registerRoute(
    _method: HttpMethod,
    _path: string,
    handler: HttpRouteHandler
  ): void {
    this.handler = handler;
  }

  async dispatch(request: IHttpServerRequest): Promise<IHttpServerResponse> {
    if (!this.handler) {
      return { status: 500, body: { error: "No handler registered" } };
    }
    return this.handler(request);
  }
}
