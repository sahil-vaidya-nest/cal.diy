import { Module } from "@nestjs/common";

import { UsersModule } from "@/modules/users/users.module";

import { DjIntegrationController } from "./dj-integration.controller";
import { DjIntegrationService } from "./dj-integration.service";
import { PrismaModule } from "@/modules/prisma/prisma.module";
import { DjIntegrationGuard } from "@/modules/dj-integration/dj-integration.guard";
import { DjMicrosoftCallbackController } from "@/modules/dj-integration/dj-microsoft-callback.controller";

@Module({
  imports: [UsersModule, PrismaModule,],
  controllers: [DjIntegrationController,  DjMicrosoftCallbackController,],
  providers: [DjIntegrationService,DjIntegrationGuard],
})
export class DjIntegrationModule {}