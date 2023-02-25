import fs from "fs";

type Body = {
  [key: string]: string;
};

export const getBodyFromFs = async () => {
  const body = await fs.promises.readFile("stored-files/lastReqBody.json", {
    encoding: "utf-8",
  });

  return JSON.parse(body);
};
