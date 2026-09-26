import {
  Controller,
  Get,
  Query,
  BadRequestException,
} from "@nestjs/common";

import { DjIntegrationService } from "./dj-integration.service";

@Controller("dj-integration/microsoft")
export class DjMicrosoftCallbackController {
  constructor(
    private readonly djIntegrationService: DjIntegrationService
  ) {}

  @Get("callback")
  async callback(
    @Query("code") code?: string,
    @Query("state") state?: string
  ) {
    if (!code || !state) {
      throw new BadRequestException(
        "Microsoft authorization code or state is missing."
      );
    }

    return this.djIntegrationService.handleMicrosoftCallback(
      code,
      state
    );
  }
}