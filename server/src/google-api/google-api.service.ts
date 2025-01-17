import { ConflictException, ForbiddenException, HttpStatus, Inject, Injectable, Logger, NotFoundException } from '@nestjs/common';
import { ConfigType } from '@nestjs/config';
import to from 'await-to-js';
import { GaxiosError, GaxiosResponse } from 'gaxios';
import { OAuth2Client } from 'google-auth-library';
import { admin_directory_v1, calendar_v3, google, people_v1 } from 'googleapis';
import appConfig from 'src/config/env/app.config';
import { GoogleAPIErrorMapper } from 'src/helpers/google-api-error.mapper';
import { OAuthTokenResponse } from '../auth/dto';
import { IGoogleApiService } from './interfaces/google-api.interface';

@Injectable()
export class GoogleApiService implements IGoogleApiService {
  constructor(
    @Inject(appConfig.KEY) private config: ConfigType<typeof appConfig>,
    private logger: Logger,
  ) { }
  getOAuthClient(): OAuth2Client {
    return new google.auth.OAuth2(this.config.oAuthClientId, this.config.oAuthClientSecret, this.config.oAuthRedirectUrl);
  }

  getOAuthUrl(client: 'web' | 'chrome') {
    const scopes = [
      'https://www.googleapis.com/auth/admin.directory.resource.calendar.readonly',
      'https://apps-apis.google.com/a/feeds/groups/',
      'https://www.googleapis.com/auth/calendar',
      'https://www.googleapis.com/auth/userinfo.email',
      'https://www.googleapis.com/auth/userinfo.profile',
      'https://www.googleapis.com/auth/directory.readonly',
      'https://www.googleapis.com/auth/admin.directory.group.readonly',
      'https://www.googleapis.com/auth/admin.directory.group.member.readonly',
    ];

    const oAuthClient = this.getOAuthClient();
    const url = oAuthClient.generateAuthUrl({
      access_type: 'offline',
      scope: scopes,
      response_type: 'code',
      state: client,
    });

    return url;
  }

  async getToken(oauth2Client: OAuth2Client, code: string): Promise<OAuthTokenResponse> {
    const [err, response]: [GaxiosError, OAuthTokenResponse] = await to(oauth2Client.getToken(code));

    if (err) {
      GoogleAPIErrorMapper.handleError(err);
    }

    oauth2Client.setCredentials(response.tokens);

    return response as OAuthTokenResponse;
  }

  async getCalendarResources(oauth2Client: OAuth2Client) {
    const service = google.admin({ version: 'directory_v1', auth: oauth2Client });
    const options = { customer: 'my_customer' };

    const [err, res]: [GaxiosError, GaxiosResponse<admin_directory_v1.Schema$CalendarResources>] = await to(service.resources.calendars.list(options));

    if (err) {
      GoogleAPIErrorMapper.handleError(err, (status: HttpStatus) => {
        if (status === HttpStatus.NOT_FOUND) {
          throw new NotFoundException('No directory resources found. Are you using an organization account?');
        }
      });
    }

    if (res.status !== 200) {
      throw new NotFoundException("Couldn't obtain directory resources");
    }

    return res.data;
  }

  async createCalenderEvent(oauth2Client: OAuth2Client, event: calendar_v3.Schema$Event) {
    const calendar = google.calendar({ version: 'v3', auth: oauth2Client });

    const [err, result]: [GaxiosError, GaxiosResponse<calendar_v3.Schema$Event>] = await to(
      calendar.events.insert({
        calendarId: 'primary',
        conferenceDataVersion: 1,
        sendUpdates: 'all',
        sendNotifications: true,
        requestBody: {
          ...event,
        },
      }),
    );

    if (err) {
      GoogleAPIErrorMapper.handleError(err);
    }

    if (result.status !== 200) {
      throw new ConflictException("Couldn't book room. Please try again later.");
    }

    return result.data;
  }

  async getCalenderSchedule(
    oauth2Client: OAuth2Client,
    start: string,
    end: string,
    timeZone: string,
    rooms: string[],
  ): Promise<{
    [key: string]: calendar_v3.Schema$FreeBusyCalendar;
  }> {
    const calendar = google.calendar({ version: 'v3', auth: oauth2Client });
    const [err, roomsFreeBusy]: [GaxiosError, GaxiosResponse<calendar_v3.Schema$FreeBusyResponse>] = await to(
      calendar.freebusy.query({
        requestBody: {
          timeMin: start,
          timeMax: end,
          timeZone,
          calendarExpansionMax: 100,
          items: rooms.map((email) => {
            return {
              id: email,
            };
          }),
        },
      }),
    );

    if (err) {
      GoogleAPIErrorMapper.handleError(err);
    }

    return roomsFreeBusy.data.calendars || {};
  }

  async getCalenderEvent(oauth2Client: OAuth2Client, id: string): Promise<calendar_v3.Schema$Event> {
    const calendar = google.calendar({ version: 'v3', auth: oauth2Client });

    const [err, res]: [GaxiosError, GaxiosResponse<calendar_v3.Schema$Event>] = await to(
      calendar.events.get({
        eventId: id,
        calendarId: 'primary',
      }),
    );

    if (err) {
      GoogleAPIErrorMapper.handleError(err);
    }

    return res.data;
  }

  async getCalenderEvents(oauth2Client: OAuth2Client, start: string, end: string, timeZone: string, limit: number = 30): Promise<calendar_v3.Schema$Event[]> {
    const calendar = google.calendar({ version: 'v3', auth: oauth2Client });

    // https://developers.google.com/calendar/api/v3/reference/events/list
    const eventTypes = ['workingLocation', 'default', 'fromGmail'];
    const [err, result]: [GaxiosError, GaxiosResponse<calendar_v3.Schema$Events>] = await to(
      calendar.events.list({
        calendarId: 'primary',
        timeMin: start,
        timeMax: end,
        timeZone,
        eventTypes,
        maxResults: limit,
        singleEvents: true,
        orderBy: 'startTime',
      }),
    );

    if (err) {
      GoogleAPIErrorMapper.handleError(err);
    }

    return result.data.items;
  }

  async updateCalenderEvent(oauth2Client: OAuth2Client, id: string, event: calendar_v3.Schema$Event): Promise<calendar_v3.Schema$Event> {
    const calendar = google.calendar({ version: 'v3', auth: oauth2Client });

    const [err, res]: [GaxiosError, GaxiosResponse<calendar_v3.Schema$Event>] = await to(
      calendar.events.update({
        eventId: id,
        calendarId: 'primary',
        requestBody: event,
        sendNotifications: true,
        sendUpdates: 'all',
        conferenceDataVersion: 1,
      }),
    );

    if (err) {
      GoogleAPIErrorMapper.handleError(err);
    }

    if (res.status !== 200) {
      throw new ForbiddenException('Could not update the event at this moment');
    }

    return res.data;
  }

  async deleteEvent(oauth2Client: OAuth2Client, id: string): Promise<void> {
    const calendar = google.calendar({ version: 'v3', auth: oauth2Client });
    const [err, _]: [GaxiosError, GaxiosResponse<void>] = await to(
      calendar.events.delete({
        calendarId: 'primary',
        eventId: id,

        sendUpdates: 'all',
        sendNotifications: true,
      }),
    );

    if (err) {
      GoogleAPIErrorMapper.handleError(err);
    }
  }

  //https://admin.googleapis.com/admin/directory/v1/groups/{groupKey}/members
  async searchGroups(oauth2Client: OAuth2Client, groupKey: string): Promise<admin_directory_v1.Schema$Members> {
    const service = google.admin({ version: 'directory_v1', auth: oauth2Client });
    const options = { groupKey };
    console.log('Calling service.members.list with options:', options);
    const [err, res]: [GaxiosError, GaxiosResponse<admin_directory_v1.Schema$Members>] = await to(service.members.list(options));
    if (err) {
      console.error('Error calling service.members.list:', err);
      GoogleAPIErrorMapper.handleError(err);
      return { members: [] };
    }

    if (!res || !res.data) {
      console.error('No response data received from service.members.list');
      return { members: [] };
    }
    console.log('Search Groups Response:', res.data);
    return res.data || { members: [] };
  }
  // https://developers.google.com/people/api/rest/v1/people/searchDirectoryPeople
  async searchPeople(oauth2Client: OAuth2Client, query: string): Promise<people_v1.Schema$Person[]> {
    const peopleService = google.people({ version: 'v1', auth: oauth2Client });

    const [err, res]: [GaxiosError, GaxiosResponse<people_v1.Schema$SearchDirectoryPeopleResponse>] = await to(
      peopleService.people.searchDirectoryPeople({
        query,
        readMask: 'emailAddresses',
        pageSize: 10,
        sources: ['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'],
      }),
    );
    console.log('Search People Response:', res.data);

    if (err) {
      this.logger.error("Couldn't search directory people: ", err);
      return [];
    }

    return res.data?.people || [];
  }
}
