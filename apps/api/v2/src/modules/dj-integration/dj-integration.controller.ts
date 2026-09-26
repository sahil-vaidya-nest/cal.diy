// import { Body, Controller, Post, UseGuards } from "@nestjs/common";
// import {
//   IsEmail,
//   IsOptional,
//   IsString,
//   IsUUID,
// } from "class-validator";

// import { DjIntegrationService } from "./dj-integration.service";
// import { DjIntegrationGuard } from "@/modules/dj-integration/dj-integration.guard";

// class FindDjAdvisorDto {
//   @IsEmail()
//   email!: string;
// }

// class ProvisionDjAdvisorDto {
//   @IsEmail()
//   email!: string;

//   @IsString()
//   name!: string;

//   @IsUUID()
//   djUserId!: string;

//   @IsOptional()
//   @IsString()
//   timeZone?: string;
// }
//   class MicrosoftConnectDto {
//   @IsString()
//   state!: string;
// }

// class BusyTimesDto {
//   @IsString()
//   externalUserId!: string;

//   @IsString()
//   dateFrom!: string;

//   @IsString()
//   dateTo!: string;

//   @IsString()
//   timeZone!: string;
// }
// class CreateDjBookingDto {
//   @IsString()
//   externalUserId!: string;

//   @IsString()
//   appointmentId!: string;

//   @IsString()
//   meetingNumber!: string;

//   @IsString()
//   title!: string;

//   @IsOptional()
//   @IsString()
//   description?: string | null;

//   @IsString()
//   start!: string;

//   @IsString()
//   end!: string;

//   @IsString()
//   timeZone!: string;

//   @IsString()
//   meetingMode!: string;

//   attendee!: {
//     name: string;
//     email: string;
//   };
// }
// @Controller("dj-integration")
// @UseGuards(DjIntegrationGuard)
// export class DjIntegrationController {
//   constructor(
//     private readonly djIntegrationService: DjIntegrationService
//   ) {}

//   @Post("advisor/find")
//   async findAdvisor(@Body() body: FindDjAdvisorDto) {
//     const user = await this.djIntegrationService.findAdvisor(body.email);

//     return {
//       data: user,
//     };
//   }

//   @Post("advisor/provision")
//   async provisionAdvisor(@Body() body: ProvisionDjAdvisorDto) {
//     const result =
//       await this.djIntegrationService.provisionAdvisor(body);

//     return {
//       data: result,
//     };
//   }


// @Post("microsoft/connect")
// async connectMicrosoft(@Body() body: MicrosoftConnectDto) {
//   const result =
//     await this.djIntegrationService.getMicrosoftConnectUrl(body.state);

//   return {
//     data: result,
//   };
// }
// @Post("calendar/busy-times")
// async getBusyTimes(@Body() body: BusyTimesDto) {
//   return this.djIntegrationService.getBusyTimes(body);
// }
// @Post("calendar/book")
// async createBooking(@Body() body: CreateDjBookingDto) {
//   return this.djIntegrationService.createBooking(body);
// }
// }
import { Body, Controller, Post, UseGuards } from "@nestjs/common";
import {
  IsEmail,
  IsOptional,
  IsString,
  IsUUID,
  ValidateNested,
} from "class-validator";
import { Type } from "class-transformer";

import { DjIntegrationService } from "./dj-integration.service";
import { DjIntegrationGuard } from "@/modules/dj-integration/dj-integration.guard";

class FindDjAdvisorDto {
  @IsEmail()
  email!: string;
}

class ProvisionDjAdvisorDto {
  @IsEmail()
  email!: string;

  @IsString()
  name!: string;

  @IsUUID()
  djUserId!: string;

  @IsOptional()
  @IsString()
  timeZone?: string;
}

class MicrosoftConnectDto {
  @IsString()
  state!: string;
}

class BusyTimesDto {
  @IsString()
  externalUserId!: string;

  @IsString()
  dateFrom!: string;

  @IsString()
  dateTo!: string;

  @IsString()
  timeZone!: string;
}

// NEW
class DjBookingAttendeeDto {
  @IsString()
  name!: string;

  @IsEmail()
  email!: string;
}

class CreateDjBookingDto {
  @IsString()
  externalUserId!: string;

  @IsString()
  appointmentId!: string;

  @IsString()
  meetingNumber!: string;

  @IsString()
  title!: string;

  @IsOptional()
  @IsString()
  description?: string | null;

  @IsString()
  start!: string;

  @IsString()
  end!: string;

  @IsString()
  timeZone!: string;

  @IsString()
  meetingMode!: string;

  // FIXED
  @ValidateNested()
  @Type(() => DjBookingAttendeeDto)
  attendee!: DjBookingAttendeeDto;
}

@Controller("dj-integration")
@UseGuards(DjIntegrationGuard)
export class DjIntegrationController {
  constructor(
    private readonly djIntegrationService: DjIntegrationService
  ) {}

  @Post("advisor/find")
  async findAdvisor(@Body() body: FindDjAdvisorDto) {
    const user =
      await this.djIntegrationService.findAdvisor(body.email);

    return {
      data: user,
    };
  }

  @Post("advisor/provision")
  async provisionAdvisor(@Body() body: ProvisionDjAdvisorDto) {
    const result =
      await this.djIntegrationService.provisionAdvisor(body);

    return {
      data: result,
    };
  }

  @Post("microsoft/connect")
  async connectMicrosoft(@Body() body: MicrosoftConnectDto) {
    const result =
      await this.djIntegrationService.getMicrosoftConnectUrl(
        body.state
      );

    return {
      data: result,
    };
  }

  @Post("calendar/busy-times")
  async getBusyTimes(@Body() body: BusyTimesDto) {
    return this.djIntegrationService.getBusyTimes(body);
  }

  @Post("calendar/book")
  async createBooking(@Body() body: CreateDjBookingDto) {
    return this.djIntegrationService.createBooking(body);
  }
}