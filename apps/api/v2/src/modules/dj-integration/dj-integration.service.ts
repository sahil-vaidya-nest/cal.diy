import { BadRequestException,Injectable,UnauthorizedException } from "@nestjs/common";



import { jwtVerify } from "jose";
import { PrismaWriteService } from "@/modules/prisma/prisma-write.service";
import { UsersRepository } from "@/modules/users/users.repository";


@Injectable()
export class DjIntegrationService {
  constructor(
    private readonly usersRepository: UsersRepository,
    private readonly dbWrite: PrismaWriteService
  ) {}

  async findAdvisor(email: string) {
    const normalizedEmail = email.trim().toLowerCase();

    const user = await this.usersRepository.findByEmail(normalizedEmail);

    if (!user) {
      return null;
    }

    return {
      id: user.id,
      email: user.email,
      username: user.username,
    };
  }

  async provisionAdvisor(input: {
    email: string;
    name: string;
    djUserId: string;
    timeZone?: string;
  }) {
    const email = input.email.trim().toLowerCase();

    // Idempotent: existing Cal.diy user reuse karo.
    const existingUser = await this.usersRepository.findByEmail(email);

    if (existingUser) {
      return {
        created: false,
        user: {
          id: existingUser.id,
          email: existingUser.email,
          username: existingUser.username,
        },
      };
    }

    const username = this.buildUsername(email, input.djUserId);

    const user = await this.dbWrite.prisma.user.create({
      data: {
        email,
        name: input.name.trim(),
        username,
        timeZone: input.timeZone ?? "Asia/Kolkata",

        // Hidden/internal Cal.diy user
        completedOnboarding: true,
        isPlatformManaged: true,
    

        // DJ mapping
        metadata: {
          djUserId: input.djUserId,
          managedBy: "DJ_IMMIGRATION",
        },
      },
      select: {
        id: true,
        email: true,
        username: true,
      },
    });

    return {
      created: true,
      user,
    };
  }

  private buildUsername(email: string, djUserId: string): string {
    const emailPrefix =
      email
        .split("@")[0]
        ?.toLowerCase()
        .replace(/[^a-z0-9]/g, "") || "advisor";

    const suffix = djUserId.replace(/-/g, "").slice(0, 8).toLowerCase();

    return `${emailPrefix}-${suffix}`;
  }
async getMicrosoftConnectUrl(state: string) {
  const app = await this.dbWrite.prisma.app.findUnique({
    where: {
      slug: "msteams",
    },
    select: {
      keys: true,
    },
  });

  const appKeys =
    app?.keys && typeof app.keys === "object" && !Array.isArray(app.keys)
      ? app.keys
      : null;

  const clientId =
    appKeys &&
    "client_id" in appKeys &&
    typeof appKeys.client_id === "string"
      ? appKeys.client_id
      : null;

  if (!clientId) {
    throw new BadRequestException(
      "Microsoft Teams client_id is not configured."
    );
  }

  // const calWebUrl = process.env.NEXT_PUBLIC_WEBAPP_URL;

  // if (!calWebUrl) {
  //   throw new BadRequestException(
  //     "Cal.diy web URL is not configured."
  //   );
  // }

  // const redirectUri =
  //   `${calWebUrl.replace(/\/$/, "")}/api/dj-integration/microsoft/callback`;
const apiUrl = process.env.DJ_INTEGRATION_API_URL;

if (!apiUrl) {
  throw new BadRequestException(
    "DJ integration API URL is not configured."
  );
}


const redirectUri =
  `${apiUrl.replace(/\/$/, "")}/dj-integration/microsoft/callback`;
const scopes = [
  "openid",
  "profile",
  "email",
  "User.Read",
  "Calendars.Read",
  "Calendars.ReadWrite",
  "OnlineMeetings.ReadWrite",
  "offline_access",
].join(" ");

  const params = new URLSearchParams({
    response_type: "code",
    scope: scopes,
    client_id: clientId,
    redirect_uri: redirectUri,
    state,
  });

  return {
    authorizationUrl:
      `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${params.toString()}`,
  };
}
async handleMicrosoftCallback(code: string, state: string) {
  // =====================================================
  // 1. VERIFY SIGNED STATE
  // =====================================================
  const stateSecret = process.env.DJ_OAUTH_STATE_SECRET;

  if (!stateSecret) {
    throw new BadRequestException(
      "DJ OAuth state secret is not configured."
    );
  }

  let payload: {
    sub?: string;
    externalUserId?: string;
    purpose?: string;
  };

  try {
    const verified = await jwtVerify(
      state,
      new TextEncoder().encode(stateSecret)
    );

    payload = verified.payload as typeof payload;
  } catch {
    throw new UnauthorizedException(
      "Invalid or expired Microsoft OAuth state."
    );
  }

  if (
    payload.purpose !== "MICROSOFT_CALENDAR_CONNECT" ||
    !payload.sub ||
    !payload.externalUserId
  ) {
    throw new UnauthorizedException(
      "Invalid Microsoft OAuth state."
    );
  }

  const calDiyUserId = Number(payload.externalUserId);

  if (!Number.isInteger(calDiyUserId) || calDiyUserId <= 0) {
    throw new UnauthorizedException(
      "Invalid Cal.diy user mapping."
    );
  }

  // =====================================================
  // 2. VERIFY CAL.DIY USER
  // =====================================================
  const calDiyUser =
    await this.dbWrite.prisma.user.findUnique({
      where: {
        id: calDiyUserId,
      },
      select: {
        id: true,
        email: true,
      },
    });

  if (!calDiyUser) {
    throw new BadRequestException(
      "Mapped Cal.diy user does not exist."
    );
  }

  // =====================================================
  // 3. GET MICROSOFT APP CREDENTIALS
  // =====================================================
  const app = await this.dbWrite.prisma.app.findUnique({
    where: {
      slug: "msteams",
    },
    select: {
      keys: true,
    },
  });

  const appKeys =
    app?.keys &&
    typeof app.keys === "object" &&
    !Array.isArray(app.keys)
      ? app.keys
      : null;

  const clientId =
    appKeys &&
    "client_id" in appKeys &&
    typeof appKeys.client_id === "string"
      ? appKeys.client_id
      : null;

  const clientSecret =
    appKeys &&
    "client_secret" in appKeys &&
    typeof appKeys.client_secret === "string"
      ? appKeys.client_secret
      : null;

  if (!clientId || !clientSecret) {
    throw new BadRequestException(
      "Microsoft credentials are not configured."
    );
  }

  const apiUrl = process.env.DJ_INTEGRATION_API_URL;

  if (!apiUrl) {
    throw new BadRequestException(
      "DJ integration API URL is not configured."
    );
  }

  const redirectUri =
    `${apiUrl.replace(/\/$/, "")}/dj-integration/microsoft/callback`;

  // =====================================================
  // 4. EXCHANGE MICROSOFT AUTHORIZATION CODE
  // =====================================================
  const microsoftScopes =
    "openid profile email User.Read Calendars.Read Calendars.ReadWrite OnlineMeetings.ReadWrite offline_access";

  const tokenBody = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: "authorization_code",
    code,
    scope: microsoftScopes,
    redirect_uri: redirectUri,
  });

  const tokenResponse = await fetch(
    "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    {
      method: "POST",
      headers: {
        "Content-Type":
          "application/x-www-form-urlencoded;charset=UTF-8",
      },
      body: tokenBody.toString(),
    }
  );

  const tokenData = await tokenResponse.json();

  if (!tokenResponse.ok) {
    throw new BadRequestException(
      "Unable to exchange Microsoft authorization code."
    );
  }

  // =====================================================
  // 5. GET MICROSOFT ACCOUNT
  // =====================================================
  const graphResponse = await fetch(
    "https://graph.microsoft.com/v1.0/me",
    {
      headers: {
        Authorization: `Bearer ${tokenData.access_token}`,
      },
    }
  );

  const graphUser = await graphResponse.json();

  if (!graphResponse.ok) {
    throw new BadRequestException(
      "Unable to fetch Microsoft account."
    );
  }

  tokenData.email =
    graphUser.mail ??
    graphUser.userPrincipalName ??
    null;

  tokenData.expiry_date =
    Math.round(Date.now() / 1000) +
    Number(tokenData.expires_in ?? 0);

  delete tokenData.expires_in;

  // =====================================================
  // 6. REMOVE ONLY THIS USER'S OLD MICROSOFT CREDENTIALS
  // =====================================================
  await this.dbWrite.prisma.credential.deleteMany({
    where: {
      userId: calDiyUserId,
      OR: [
        {
          type: "office365_calendar",
          appId: "office365-calendar",
        },
        {
          type: "office365_video",
          appId: "msteams",
        },
      ],
    },
  });

  // =====================================================
  // 7. CREATE OUTLOOK CALENDAR CREDENTIAL
  // Used by C6 / busy-times
  // =====================================================
  const calendarCredential =
    await this.dbWrite.prisma.credential.create({
      data: {
        type: "office365_calendar",
        key: tokenData,
        userId: calDiyUserId,
        appId: "office365-calendar",
      },
      select: {
        id: true,
      },
    });

  // =====================================================
  // 8. CREATE TEAMS VIDEO CREDENTIAL
  // Used by A8 / Teams meeting
  // =====================================================
  const teamsCredential =
    await this.dbWrite.prisma.credential.create({
      data: {
        type: "office365_video",
        key: tokenData,
        userId: calDiyUserId,
        appId: "msteams",
      },
      select: {
        id: true,
      },
    });

  // =====================================================
  // 9. UPDATE DJ BACKEND
  // IMPORTANT:
  // externalCredentialId = CALENDAR credential
  // =====================================================
  const djBackendUrl = process.env.DJ_BACKEND_URL;
  const integrationSecret =
    process.env.DJ_INTEGRATION_SECRET;

  if (!djBackendUrl || !integrationSecret) {
    throw new BadRequestException(
      "DJ backend integration configuration is missing."
    );
  }

  const djResponse = await fetch(
    `${djBackendUrl.replace(
      /\/$/,
      ""
    )}/api/v1/calendar-integrations/internal/microsoft-connected`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-dj-integration-secret": integrationSecret,
      },
      body: JSON.stringify({
        userId: payload.sub,
        externalUserId: String(calDiyUserId),

        // C6 needs calendar credential
        externalCredentialId: String(
          calendarCredential.id
        ),

        email: tokenData.email,
      }),
    }
  );

  if (!djResponse.ok) {
    throw new BadRequestException(
      "Microsoft credentials created, but DJ calendar integration update failed."
    );
  }

  // =====================================================
  // 10. SAFE RESPONSE - NEVER RETURN TOKENS
  // =====================================================
  return {
    connected: true,
    djUserId: payload.sub,
    externalUserId: String(calDiyUserId),
    externalCredentialId: String(
      calendarCredential.id
    ),
    teamsCredentialConfigured:
      Boolean(teamsCredential.id),
    email: tokenData.email,
  };
}
async getBusyTimes(input: {
  externalUserId: string;
  dateFrom: string;
  dateTo: string;
  timeZone: string;
}) {
  const userId = Number(input.externalUserId);

  if (!Number.isInteger(userId) || userId <= 0) {
    throw new BadRequestException("Invalid Cal.diy user mapping.");
  }

  // 1. Get advisor Outlook credential
  const credential = await this.dbWrite.prisma.credential.findFirst({
    where: {
      userId,
      type: "office365_calendar",
      appId: "office365-calendar",
    },
    orderBy: { id: "desc" },
    select: {
      id: true,
      key: true,
    },
  });

  if (!credential) {
    throw new BadRequestException(
      "Advisor Outlook calendar credential not found."
    );
  }

  const key = credential.key as Record<string, any>;

  if (typeof key?.access_token !== "string") {
    throw new BadRequestException(
      "Advisor Outlook access token not found."
    );
  }

  let accessToken: string = key.access_token;

  // 2. Microsoft Graph helper
  const callGraph = (token: string) => {
    const query = new URLSearchParams({
      startDateTime: input.dateFrom,
      endDateTime: input.dateTo,
      $select: "start,end,showAs,isCancelled",
    });

    return fetch(
      `https://graph.microsoft.com/v1.0/me/calendarView?${query.toString()}`,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          // Prefer: `outlook.timezone="${input.timeZone}"`,
        },
      }
    );
  };

  // 3. First Graph request
  let response = await callGraph(accessToken);

  // 4. Token expired → refresh → retry once
  if (response.status === 401) {
    if (typeof key.refresh_token !== "string") {
      throw new BadRequestException(
        "Microsoft session expired. Please reconnect Outlook."
      );
    }

    const app = await this.dbWrite.prisma.app.findUnique({
      where: {
        slug: "msteams",
      },
      select: {
        keys: true,
      },
    });

    const appKeys = app?.keys as Record<string, any> | undefined;

    if (
      typeof appKeys?.client_id !== "string" ||
      typeof appKeys?.client_secret !== "string"
    ) {
      throw new BadRequestException(
        "Microsoft app credentials are not configured."
      );
    }

    const tokenResponse = await fetch(
      "https://login.microsoftonline.com/common/oauth2/v2.0/token",
      {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: new URLSearchParams({
          client_id: appKeys.client_id,
          client_secret: appKeys.client_secret,
          grant_type: "refresh_token",
          refresh_token: key.refresh_token,
          scope:
            "openid profile email User.Read Calendars.Read Calendars.ReadWrite OnlineMeetings.ReadWrite offline_access",
        }),
      }
    );

    const tokenData = await tokenResponse.json();

    if (
      !tokenResponse.ok ||
      typeof tokenData?.access_token !== "string"
    ) {
      console.error("MICROSOFT TOKEN REFRESH ERROR:", {
        status: tokenResponse.status,
        error: tokenData?.error,
        message: tokenData?.error_description,
      });

      throw new BadRequestException(
        "Microsoft session expired. Please reconnect Outlook."
      );
    }

    // Guaranteed string
    const newAccessToken: string = tokenData.access_token;
    accessToken = newAccessToken;

    // 5. Save refreshed token
    await this.dbWrite.prisma.credential.update({
      where: {
        id: credential.id,
      },
      data: {
        key: {
          ...key,
          ...tokenData,
          access_token: newAccessToken,
          refresh_token:
            typeof tokenData.refresh_token === "string"
              ? tokenData.refresh_token
              : key.refresh_token,
          expiry_date:
            Date.now() +
            (typeof tokenData.expires_in === "number"
              ? tokenData.expires_in
              : 3600) *
              1000,
        },
      },
    });

    // 6. Retry Graph once
    response = await callGraph(newAccessToken);
  }

  // 7. Process Graph response
  const data = await response.json();

  if (!response.ok) {
    console.error(
      "MICROSOFT GRAPH BUSY TIME ERROR:",
      response.status,
      data
    );

    throw new BadRequestException(
      "Unable to fetch advisor Outlook busy times."
    );
  }

  // 8. Return only busy events
  const busyTimes = (data.value ?? [])
    .filter(
      (event: any) =>
        !event.isCancelled &&
        event.showAs !== "free" &&
        event.start?.dateTime &&
        event.end?.dateTime
    )
    .map((event: any) => ({
      start: event.start.dateTime,
      end: event.end.dateTime,
    }));
console.log("OUTLOOK BUSY TIMES:", busyTimes);
  return {
    data: busyTimes,
  };
}
async createBooking(input: {
  externalUserId: string;
  appointmentId: string;
  meetingNumber: string;
  title: string;
  description?: string | null;
  start: string;
  end: string;
  timeZone: string;
  meetingMode: string;
  attendee: {
    name: string;
    email: string;
  };
}) {
  console.log("DJ BOOKING INPUT:", JSON.stringify(input, null, 2));
  if (
  !input.attendee ||
  !input.attendee.email ||
  !input.attendee.name
) {
  throw new BadRequestException(
    "Meeting attendee name and email are required."
  );
}
  const userId = Number(input.externalUserId);
 
  if (!Number.isInteger(userId) || userId <= 0) {
    throw new BadRequestException("Invalid Cal.diy user mapping.");
  }
 
  // =====================================================
  // 1. GET ADVISOR MICROSOFT CALENDAR CREDENTIAL
  // =====================================================
 
  const credential =
    await this.dbWrite.prisma.credential.findFirst({
      where: {
        userId,
        type: "office365_calendar",
        appId: "office365-calendar",
      },
      orderBy: {
        id: "desc",
      },
      select: {
        id: true,
        key: true,
      },
    });
 
  if (!credential) {
    throw new BadRequestException(
      "Advisor Outlook calendar credential not found."
    );
  }
 
  const key = credential.key as Record<string, any>;
 
  if (typeof key?.access_token !== "string") {
    throw new BadRequestException(
      "Advisor Outlook access token not found."
    );
  }
 
  let accessToken: string = key.access_token;
const toLocalDateTime = (
  value: string,
  timeZone: string
): string => {
  const date = new Date(value);
 
  if (Number.isNaN(date.getTime())) {
    throw new BadRequestException(
      "Invalid meeting date/time."
    );
  }
 
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hourCycle: "h23",
  }).formatToParts(date);
 
  const get = (type: Intl.DateTimeFormatPartTypes) =>
    parts.find((part) => part.type === type)?.value;
 
  return `${get("year")}-${get("month")}-${get("day")}T${get("hour")}:${get("minute")}:${get("second")}`;
};
 
const graphStart = toLocalDateTime(
  input.start,
  input.timeZone
);
 
const graphEnd = toLocalDateTime(
  input.end,
  input.timeZone
);
  // =====================================================
  // 2. MICROSOFT GRAPH BOOKING
  // =====================================================
 
  const callGraph = (token: string) => {
    const isOnline = input.meetingMode === "ONLINE";
 
    return fetch(
      "https://graph.microsoft.com/v1.0/me/events",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          subject: input.title,
 
          body: {
            contentType: "text",
            content: input.description ?? "",
          },
 
         start: {
  dateTime: graphStart,
  timeZone: input.timeZone,
},
 
end: {
  dateTime: graphEnd,
  timeZone: input.timeZone,
},
 
          attendees: [
            {
              emailAddress: {
                address: input.attendee.email,
                name: input.attendee.name,
              },
              type: "required",
            },
          ],
 
          showAs: "busy",
 
          isOnlineMeeting: isOnline,
 
          ...(isOnline
            ? {
                onlineMeetingProvider: "teamsForBusiness",
              }
            : {}),
 
          transactionId: input.appointmentId,
        }),
      }
    );
  };
 
  // =====================================================
  // 3. FIRST ATTEMPT
  // =====================================================
 
  let response = await callGraph(accessToken);
 
  // =====================================================
  // 4. TOKEN EXPIRED -> REFRESH -> RETRY ONCE
  // =====================================================
 
  if (response.status === 401) {
    if (typeof key.refresh_token !== "string") {
      throw new BadRequestException(
        "Microsoft session expired. Please reconnect Outlook."
      );
    }
 
    const app = await this.dbWrite.prisma.app.findUnique({
      where: {
        slug: "msteams",
      },
      select: {
        keys: true,
      },
    });
 
    const appKeys =
      app?.keys as Record<string, any> | undefined;
 
    if (
      typeof appKeys?.client_id !== "string" ||
      typeof appKeys?.client_secret !== "string"
    ) {
      throw new BadRequestException(
        "Microsoft app credentials are not configured."
      );
    }
 
    const tokenResponse = await fetch(
      "https://login.microsoftonline.com/common/oauth2/v2.0/token",
      {
        method: "POST",
        headers: {
          "Content-Type":
            "application/x-www-form-urlencoded",
        },
        body: new URLSearchParams({
          client_id: appKeys.client_id,
          client_secret: appKeys.client_secret,
          grant_type: "refresh_token",
          refresh_token: key.refresh_token,
          scope:
            "openid profile email User.Read Calendars.Read Calendars.ReadWrite OnlineMeetings.ReadWrite offline_access",
        }),
      }
    );
 
    const tokenData = await tokenResponse.json();
 
    if (
      !tokenResponse.ok ||
      typeof tokenData?.access_token !== "string"
    ) {
      throw new BadRequestException(
        "Microsoft session expired. Please reconnect Outlook."
      );
    }
 
    const newAccessToken: string =
      tokenData.access_token;
 
    accessToken = newAccessToken;
 
    await this.dbWrite.prisma.credential.update({
      where: {
        id: credential.id,
      },
      data: {
        key: {
          ...key,
          ...tokenData,
 
          access_token: newAccessToken,
 
          refresh_token:
            typeof tokenData.refresh_token === "string"
              ? tokenData.refresh_token
              : key.refresh_token,
 
          expiry_date:
            Date.now() +
            (typeof tokenData.expires_in === "number"
              ? tokenData.expires_in
              : 3600) *
              1000,
        },
      },
    });
 
    response = await callGraph(newAccessToken);
  }
 
  // =====================================================
  // 5. PROCESS CREATED OUTLOOK EVENT
  // =====================================================
 
  const event = await response.json();
 
  if (!response.ok) {
    console.error(
      "MICROSOFT GRAPH BOOKING ERROR:",
      response.status,
      event
    );
 
    throw new BadRequestException(
      "Unable to create advisor Outlook meeting."
    );
  }
 
  // =====================================================
  // 6. RETURN SAFE BOOKING RESULT
  // =====================================================
  console.log(
  "GRAPH CREATED EVENT TIME:",
  JSON.stringify(
    {
      start: event.start,
      end: event.end,
      originalStartTimeZone: event.originalStartTimeZone,
      originalEndTimeZone: event.originalEndTimeZone,
    },
    null,
    2
  )
);
  return {
    data: {
      externalBookingId: event.id,
      externalEventId: event.id,
 
      meetingLink:
        event.onlineMeeting?.joinUrl ??
        event.onlineMeetingUrl ??
        null,
 
      start: input.start,
      end: input.end,
    },
  };
 
}
async rescheduleBooking(input: {
  externalUserId: string;
  externalEventId: string;
  start: string;
  end: string;
  timeZone: string;
}) {
  const userId = Number(input.externalUserId);

  if (!Number.isInteger(userId) || userId <= 0) {
    throw new BadRequestException(
      "Invalid Cal.diy user mapping."
    );
  }

  if (!input.externalEventId?.trim()) {
    throw new BadRequestException(
      "Outlook event ID is required."
    );
  }

  // =====================================================
  // 1. GET ADVISOR MICROSOFT CALENDAR CREDENTIAL
  // =====================================================

  const credential =
    await this.dbWrite.prisma.credential.findFirst({
      where: {
        userId,
        type: "office365_calendar",
        appId: "office365-calendar",
      },
      orderBy: {
        id: "desc",
      },
      select: {
        id: true,
        key: true,
      },
    });

  if (!credential) {
    throw new BadRequestException(
      "Advisor Outlook calendar credential not found."
    );
  }

  const key = credential.key as Record<string, any>;

  if (typeof key?.access_token !== "string") {
    throw new BadRequestException(
      "Advisor Outlook access token not found."
    );
  }

  let accessToken: string = key.access_token;

  // =====================================================
  // 2. CONVERT UTC INSTANT -> ADVISOR LOCAL DATETIME
  // =====================================================

  const toLocalDateTime = (
    value: string,
    timeZone: string
  ): string => {
    const date = new Date(value);

    if (Number.isNaN(date.getTime())) {
      throw new BadRequestException(
        "Invalid meeting date/time."
      );
    }

    const parts = new Intl.DateTimeFormat("en-CA", {
      timeZone,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hourCycle: "h23",
    }).formatToParts(date);

    const get = (
      type: Intl.DateTimeFormatPartTypes
    ) =>
      parts.find((part) => part.type === type)?.value;

    return `${get("year")}-${get("month")}-${get(
      "day"
    )}T${get("hour")}:${get("minute")}:${get(
      "second"
    )}`;
  };

  const graphStart = toLocalDateTime(
    input.start,
    input.timeZone
  );

  const graphEnd = toLocalDateTime(
    input.end,
    input.timeZone
  );

  // =====================================================
  // 3. PATCH EXISTING OUTLOOK EVENT
  // =====================================================

  const callGraph = (token: string) => {
    return fetch(
      `https://graph.microsoft.com/v1.0/me/events/${encodeURIComponent(
        input.externalEventId
      )}`,
      {
        method: "PATCH",

        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },

        body: JSON.stringify({
          start: {
            dateTime: graphStart,
            timeZone: input.timeZone,
          },

          end: {
            dateTime: graphEnd,
            timeZone: input.timeZone,
          },
        }),
      }
    );
  };

  // =====================================================
  // 4. FIRST ATTEMPT
  // =====================================================

  let response = await callGraph(accessToken);

  // =====================================================
  // 5. TOKEN EXPIRED -> REFRESH -> RETRY ONCE
  // =====================================================

  if (response.status === 401) {
    if (typeof key.refresh_token !== "string") {
      throw new BadRequestException(
        "Microsoft session expired. Please reconnect Outlook."
      );
    }

    const app =
      await this.dbWrite.prisma.app.findUnique({
        where: {
          slug: "msteams",
        },
        select: {
          keys: true,
        },
      });

    const appKeys =
      app?.keys as Record<string, any> | undefined;

    if (
      typeof appKeys?.client_id !== "string" ||
      typeof appKeys?.client_secret !== "string"
    ) {
      throw new BadRequestException(
        "Microsoft app credentials are not configured."
      );
    }

    const tokenResponse = await fetch(
      "https://login.microsoftonline.com/common/oauth2/v2.0/token",
      {
        method: "POST",

        headers: {
          "Content-Type":
            "application/x-www-form-urlencoded",
        },

        body: new URLSearchParams({
          client_id: appKeys.client_id,
          client_secret: appKeys.client_secret,

          grant_type: "refresh_token",
          refresh_token: key.refresh_token,

          scope:
            "openid profile email User.Read Calendars.Read Calendars.ReadWrite OnlineMeetings.ReadWrite offline_access",
        }),
      }
    );

    const tokenData = await tokenResponse.json();

    if (
      !tokenResponse.ok ||
      typeof tokenData?.access_token !== "string"
    ) {
      throw new BadRequestException(
        "Microsoft session expired. Please reconnect Outlook."
      );
    }

    const newAccessToken: string =
      tokenData.access_token;

    accessToken = newAccessToken;

    await this.dbWrite.prisma.credential.update({
      where: {
        id: credential.id,
      },

      data: {
        key: {
          ...key,
          ...tokenData,

          access_token: newAccessToken,

          refresh_token:
            typeof tokenData.refresh_token === "string"
              ? tokenData.refresh_token
              : key.refresh_token,

          expiry_date:
            Date.now() +
            (typeof tokenData.expires_in === "number"
              ? tokenData.expires_in
              : 3600) *
              1000,
        },
      },
    });

    response = await callGraph(newAccessToken);
  }

  // =====================================================
  // 6. PROCESS GRAPH RESPONSE
  // =====================================================

  let event: any = null;

  const responseText = await response.text();

  if (responseText) {
    try {
      event = JSON.parse(responseText);
    } catch {
      event = null;
    }
  }

  if (!response.ok) {
    console.error(
      "MICROSOFT GRAPH RESCHEDULE ERROR:",
      response.status,
      event
    );

    if (response.status === 404) {
      throw new BadRequestException(
        "Outlook meeting could not be found."
      );
    }

    throw new BadRequestException(
      "Unable to reschedule advisor Outlook meeting."
    );
  }

  // =====================================================
  // 7. RETURN SAFE RESULT
  // =====================================================

  return {
    data: {
      externalEventId: input.externalEventId,

      meetingLink:
        event?.onlineMeeting?.joinUrl ??
        event?.onlineMeetingUrl ??
        null,

      start: input.start,
      end: input.end,
    },
  };
}

async cancelBooking(input: {
  externalUserId: string;
  externalEventId: string;
  reason?: string;
}) {
  const userId = Number(input.externalUserId);

  if (!Number.isInteger(userId) || userId <= 0) {
    throw new BadRequestException(
      "Invalid Cal.diy user mapping."
    );
  }

  if (!input.externalEventId?.trim()) {
    throw new BadRequestException(
      "Outlook event ID is required."
    );
  }

  // =====================================================
  // 1. GET ADVISOR MICROSOFT CALENDAR CREDENTIAL
  // =====================================================

  const credential =
    await this.dbWrite.prisma.credential.findFirst({
      where: {
        userId,
        type: "office365_calendar",
        appId: "office365-calendar",
      },
      orderBy: {
        id: "desc",
      },
      select: {
        id: true,
        key: true,
      },
    });

  if (!credential) {
    throw new BadRequestException(
      "Advisor Outlook calendar credential not found."
    );
  }

  const key = credential.key as Record<string, any>;

  if (typeof key?.access_token !== "string") {
    throw new BadRequestException(
      "Advisor Outlook access token not found."
    );
  }

  let accessToken: string = key.access_token;

  // =====================================================
  // 2. CANCEL EXISTING OUTLOOK EVENT
  // =====================================================

  const callGraph = (token: string) => {
    return fetch(
      `https://graph.microsoft.com/v1.0/me/events/${encodeURIComponent(
        input.externalEventId
      )}/cancel`,
      {
        method: "POST",

        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },

        body: JSON.stringify({
          comment:
            input.reason?.trim() ||
            "Meeting cancelled.",
        }),
      }
    );
  };

  // =====================================================
  // 3. FIRST ATTEMPT
  // =====================================================

  let response = await callGraph(accessToken);

  // =====================================================
  // 4. TOKEN EXPIRED -> REFRESH -> RETRY ONCE
  // =====================================================

  if (response.status === 401) {
    if (typeof key.refresh_token !== "string") {
      throw new BadRequestException(
        "Microsoft session expired. Please reconnect Outlook."
      );
    }

    const app =
      await this.dbWrite.prisma.app.findUnique({
        where: {
          slug: "msteams",
        },
        select: {
          keys: true,
        },
      });

    const appKeys =
      app?.keys as Record<string, any> | undefined;

    if (
      typeof appKeys?.client_id !== "string" ||
      typeof appKeys?.client_secret !== "string"
    ) {
      throw new BadRequestException(
        "Microsoft app credentials are not configured."
      );
    }

    const tokenResponse = await fetch(
      "https://login.microsoftonline.com/common/oauth2/v2.0/token",
      {
        method: "POST",

        headers: {
          "Content-Type":
            "application/x-www-form-urlencoded",
        },

        body: new URLSearchParams({
          client_id: appKeys.client_id,
          client_secret: appKeys.client_secret,

          grant_type: "refresh_token",
          refresh_token: key.refresh_token,

          scope:
            "openid profile email User.Read Calendars.Read Calendars.ReadWrite OnlineMeetings.ReadWrite offline_access",
        }),
      }
    );

    const tokenData = await tokenResponse.json();

    if (
      !tokenResponse.ok ||
      typeof tokenData?.access_token !== "string"
    ) {
      throw new BadRequestException(
        "Microsoft session expired. Please reconnect Outlook."
      );
    }

    const newAccessToken: string =
      tokenData.access_token;

    accessToken = newAccessToken;

    await this.dbWrite.prisma.credential.update({
      where: {
        id: credential.id,
      },

      data: {
        key: {
          ...key,
          ...tokenData,

          access_token: newAccessToken,

          refresh_token:
            typeof tokenData.refresh_token === "string"
              ? tokenData.refresh_token
              : key.refresh_token,

          expiry_date:
            Date.now() +
            (typeof tokenData.expires_in === "number"
              ? tokenData.expires_in
              : 3600) *
              1000,
        },
      },
    });

    response = await callGraph(newAccessToken);
  }

  // =====================================================
  // 5. PROCESS GRAPH RESPONSE
  // =====================================================

  if (!response.ok) {
    let errorData: any = null;

    const responseText = await response.text();

    if (responseText) {
      try {
        errorData = JSON.parse(responseText);
      } catch {
        errorData = null;
      }
    }

    // Temporary while testing — remove sensitive Graph logs
    // before production.
    console.error(
      "MICROSOFT GRAPH CANCEL ERROR:",
      response.status,
      errorData
    );

    if (response.status === 404) {
      throw new BadRequestException(
        "Outlook meeting could not be found."
      );
    }

    throw new BadRequestException(
      "Unable to cancel advisor Outlook meeting."
    );
  }

  // Graph cancel normally returns 202 Accepted / empty body.

  // =====================================================
  // 6. RETURN SAFE RESULT
  // =====================================================

  return {
    data: {
      externalEventId: input.externalEventId,
      cancelled: true,
    },
  };
}
}