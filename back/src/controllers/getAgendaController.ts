import { Request, Response } from "express";
import { calendar_v3, google } from "googleapis";
import fs from "fs";

import * as dotenv from "dotenv";
dotenv.config();

export type ControllerType = {
  [key: string]: (req: Request, res: Response) => Promise<void>;
};

const getAuth = () =>
  new google.auth.GoogleAuth({
    keyFile: "./auth.json",
    scopes: [
      "https://www.googleapis.com/auth/calendar",
      "https://www.googleapis.com/auth/calendar.events",
    ],
  });

const getAgenda = () => {
  const auth = getAuth();

  const agenda = google.calendar({
    version: "v3",
    auth,
  });

  return agenda;
};

const trimEventDate = (events: calendar_v3.Schema$Event[]) => {
  return events.map((event) => ({
    start: event.start?.dateTime,
    end: event.end?.dateTime,
  }));
};

const convertEventsToSlots = (events: calendar_v3.Schema$Event[]) => {
  const SLOT_DURATION = 30 * 60 * 1000;

  const slots = trimEventDate(events)
    .map((event) => {
      const { start, end } = event;
      const slots: Date[] = [];
      if (start && end) {
        const startDate = new Date(start);

        const startZero = new Date(
          startDate.getFullYear(),
          startDate.getMonth(),
          startDate.getDate(),
          startDate.getHours(),
          startDate.getMinutes() >= 30 ? 30 : 0
        );

        let loopTime = startZero.getTime();
        const maxTime = new Date(end).getTime();

        while (loopTime < maxTime) {
          slots.push(new Date(loopTime));

          loopTime += SLOT_DURATION;
        }
      }
      return slots;
    })
    .flat();

  return slots;
};

const getAllEvents = async () => {
  const agenda = getAgenda();

  const eventList = await agenda.events.list({
    calendarId: process.env.CALENDAR_ID,
  });

  return trimEventDate(eventList.data.items || []);
};

const getSlotsByDate = async (monthYear: string) => {
  if (monthYear.match(/^\d{4}-\d{2}$/)) {
    const year = parseInt(monthYear.split("-")[0], 10);
    const month = parseInt(monthYear.split("-")[1], 10);

    const timeMin = new Date(year, month - 1, 1).toISOString();
    const timeMax = new Date(year, month, 0).toISOString();

    const agenda = getAgenda();

    const eventList = await agenda.events.list({
      calendarId: process.env.CALENDAR_ID,
      timeMin,
      timeMax,
      singleEvents: true,
    });

    return convertEventsToSlots(eventList.data.items || []);
  }
  return [];
};

const getAgendaController: ControllerType = {};

getAgendaController.getAll = async (req, res) => {
  try {
    const allEvents = await getAllEvents();
    res.send(allEvents);
  } catch (err: unknown) {
    console.error(err);
    res.sendStatus(500);
  }
};

getAgendaController.getByDate = async (req, res) => {
  try {
    const monthYear = req.body.monthYear as string;

    const slots = await getSlotsByDate(monthYear);
    res.send({ slots });
  } catch (err: unknown) {
    console.error(err);
    res.sendStatus(500);
  }
};

export default getAgendaController;
