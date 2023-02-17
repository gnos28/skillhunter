import { Request, Response } from "express";
import * as dotenv from "dotenv";
import { getAgenda } from "../utils/google";
import { exportToSheet } from "../utils/exportToSheet";
dotenv.config();

export type ControllerType = {
  [key: string]: (req: Request, res: Response) => Promise<void>;
};

type PurchasingData = {
  id: number;
  create_time: string | undefined;
  purchasedAmount: string | undefined;
  purchasedCurrency: string | undefined;
  status: string | undefined;
  payer_id: string | undefined;
  payer_name: string | undefined;
  prenom: string | undefined;
  nom: string | undefined;
  email_adress: string | undefined;
  address: string | undefined;
  date: string | undefined;
  hour: string | undefined;
  prestationName: string | undefined;
  prestationDuration: number | undefined;
  prestationPrice: number | undefined;
  telephone: string | undefined;
  instagram: string | undefined;
  whatsapp: string | undefined;
};

const storePaypalController: ControllerType = {};

storePaypalController.store = async (req, res) => {
  try {
    const purchasingData = req.body.purchasingData as PurchasingData;

    // save data to google sheet
    const nbNew = await exportToSheet([purchasingData], "840247244");
    // create new agenda event
    const purchasingDataDate = purchasingData.date;
    const splittedHour = purchasingData.hour?.split(":") || [];
    const duration = purchasingData.prestationDuration;

    if (
      nbNew === 1 &&
      purchasingDataDate &&
      splittedHour?.length > 1 &&
      duration
    ) {
      const agenda = getAgenda();

      const year = new Date(purchasingDataDate).getFullYear();
      const month = new Date(purchasingDataDate).getMonth();
      const day = new Date(purchasingDataDate).getDate();
      const hour = parseInt(splittedHour[0], 10);
      const minutes = parseInt(splittedHour[1], 10);

      const startDate = new Date(year, month, day, hour, minutes);
      console.log("startDate", startDate);

      const start = startDate.toISOString();
      console.log("start", start);
      console.log(
        "endDate",
        new Date(startDate.getTime() + duration * 60 * 1000)
      );

      const end = new Date(
        startDate.getTime() + duration * 60 * 1000
      ).toISOString();
      console.log("end", end);

      const res = await agenda.events.insert({
        calendarId: process.env.CALENDAR_ID,
        conferenceDataVersion: 0,
        sendNotifications: true,
        sendUpdates: "all",
        supportsAttachments: false,
        requestBody: {
          // request body parameters
          // {
          //   "anyoneCanAddSelf": false,
          //   "attachments": [],
          //   "attendees": [],
          //   "attendeesOmitted": false,
          //   "colorId": "my_colorId",
          //   "conferenceData": {},
          //   "created": "my_created",
          //   "creator": {},
          description: "my_description",
          end: {
            dateTime: end,
          },
          //   "endTimeUnspecified": false,
          //   "etag": "my_etag",
          //   "eventType": "my_eventType",
          //   "extendedProperties": {},
          //   "gadget": {},
          guestsCanInviteOthers: false,
          guestsCanModify: false,
          guestsCanSeeOtherGuests: false,
          //   "hangoutLink": "my_hangoutLink",
          //   "htmlLink": "my_htmlLink",
          //   "iCalUID": "my_iCalUID",
          //   "id": "my_id",
          //   "kind": "my_kind",
          //   "location": "my_location",
          //   "locked": false,
          //   "organizer": {},
          //   "originalStartTime": {},
          //   "privateCopy": false,
          //   "recurrence": [],
          //   "recurringEventId": "my_recurringEventId",
          //   "reminders": {},
          //   "sequence": 0,
          //   "source": {},
          start: {
            dateTime: start,
          },
          //   "status": "my_status",
          summary: "TITRE",
          //   "transparency": "my_transparency",
          //   "updated": "my_updated",
          //   "visibility": "my_visibility"
          // }
        },
      });

      console.log("res", res);
    }

    // send mail to customer

    // const slots = await getSlotsByDate(monthYear);
    res.send({});
  } catch (err: unknown) {
    console.error(err);
    res.sendStatus(500);
  }
};

export default storePaypalController;
