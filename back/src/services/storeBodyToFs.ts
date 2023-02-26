import fs from "fs";

type Body = {
  [key: string]: string;
};

export const storeBodyToFs = async (body: Body) => {
  console.log("storeBodyToFs", body);

  await fs.promises.writeFile(
    "stored-files/lastReqBody.json",
    JSON.stringify(body)
  );
};
