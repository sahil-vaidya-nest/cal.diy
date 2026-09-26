import {
  CanActivate,
  ExecutionContext,
  Injectable,
  UnauthorizedException,
} from "@nestjs/common";
import type { Request } from "express";
import { timingSafeEqual } from "node:crypto";

@Injectable()
export class DjIntegrationGuard implements CanActivate {
  canActivate(context: ExecutionContext): boolean {
    const request = context.switchToHttp().getRequest<Request>();

    const providedSecret = request.header("x-dj-integration-secret");
    const expectedSecret = process.env.DJ_INTEGRATION_SECRET;

    if (!providedSecret || !expectedSecret) {
      throw new UnauthorizedException("Invalid DJ integration credentials");
    }

    const providedBuffer = Buffer.from(providedSecret);
    const expectedBuffer = Buffer.from(expectedSecret);

    if (providedBuffer.length !== expectedBuffer.length) {
      throw new UnauthorizedException("Invalid DJ integration credentials");
    }

    if (!timingSafeEqual(providedBuffer, expectedBuffer)) {
      throw new UnauthorizedException("Invalid DJ integration credentials");
    }

    return true;
  }
}