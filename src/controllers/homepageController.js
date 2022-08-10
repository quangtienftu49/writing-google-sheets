import moment from "moment";
const { GoogleSpreadsheet } = require("google-spreadsheet");

const PRIVATE_KEY =
  "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDI3nVa8DNLDlab\nimjpHgZtqPCUtS5QRg7NSQFNpOAxyDxS5VeyIUS1bnWyX7/RGVtOOOfyyyDW1r8+\nhmIacXfqtkjKZFrOaHm3FWQQ2WlT0jgFA9K4IvOGASno5NUWyS1FVS4NdM/xQVZv\nzOUf9WJap1YYqyfGlCbU5dyJf27mNMGxAWXWYUCUKdmuxkzuGh+sjG2fiPOqQvP7\nu8zoKjQyzwkt5bIJb+DLQ6n9kx1JRHb9Tytr0GSjqSUX9PN9X9jDSX34FPdTR3Qr\nmAAeOmniPSf7Enlv+Xf+qbCyS7aSc4wThxhZYbzTXgtHAzDtRtMvEKTwDmu4K8uE\nLyMXydQtAgMBAAECggEACCBBwKBLfuuRUuEGTfn2Ohz1wJpUyZAP/LuJlzXGkRT/\n5b2b7tucLzfNAJ6RgLjFIhS4w3B2v2I4IeQnV1kESCVALBnjt4pv9BlkLy9QbMDq\nHb/SGDtnvSHSa3bIbQ/N6kpgHtC9Mc4aHFM+aKzRYgf3em+AIC8xo/a5lYFnf94L\npbbaHjiZ19lKRdnEIq6QfctRr1l+gxt8ltDShjkFHiLEDlK3AcxzXjsxyFaflq55\nu+qsoXdNX8g129L/AtAEFor9Z8XcznOGsDbfq7o/nnIPC6rCkxh8jCAUhlAy/0zp\ngLNm1pao2Jvu+T1EpGlaHNj++xKhw2oFVDzXDMu6EQKBgQDp3Vr1NGdVriCr2EY6\nCXfWNclXNw0yxvIEuDi7KPHpJ8PIh7TYjjdkRomxGu4scU3ixDkbyXw/RooRv0YR\nKU3kn0GUpgVj29c+6jiL9p/dZYKgRPfjFXrtibACSCAlaBwvx9XV59qy1/HUs8v+\ndX/j6/8H0pGcVmm0/Xt78DCCsQKBgQDb4ZpWjhBo0LtdkoT/5nkvSUDO8FLTfV/6\njlQ62i3z4mFKomYeKDVz7m1yz6GkBrcfzu3jZZmcY0TQ9ZSD9hFupK/tftaGz68/\nQAPa7mBpUj//o5VcplW7LxQ4aB7vq/q7XtglBK6042zVNBtHc7K0DXHBBzhRXnKf\nCNiPBrGwPQKBgQDVsTonzJaPp+ianait53Dk/4jWdKtOtpL21Q6hlixWC8vONJJ/\nPpRGwF2Ywy7W1UGB8CLuzREHEIGg7dIsZD2UpiDan0lVkdAA4SyCV/yD5PmTUPHh\nQgNtgd6edyFIjPUUg9lU9+LSgJes8A16mgseTMpgb3w2Co/Unbpz6WmqQQKBgQCJ\nGgLCNZLFyGEL13BWn76wXVyrq+35MRPHhze9+ozspRtFDj3eT/QEdYaJMC35uLY2\nfzCVuaQufzdJk9cm8SetdcK8s3nQVW9QYPoGaNx0z3RYUgev3YdXT+OryECB8RpF\n+r2LV4AYCjayOetIgjvLSRbE5VuYYOvXfgyKIgJpgQKBgAiieZLyQ2cKbmSDmF1g\nmvq2DhNuh+jVFnzJSfnix76pc+STT+Z4zNiu5f80k9qw4xbCsPuxE/Tx+/3f01po\ngbL1tGTwW4T2oF3SB9eIG75vxFhF9ARNdnZbVeH7YHkvi2UKd5FPAExwTgmVmgpp\nVMBf5nCsN7u+pOXE04SQXlZh\n-----END PRIVATE KEY-----\n";
const CLIENT_EMAIL =
  "keyin-google-sheets@lunar-analyzer-359002.iam.gserviceaccount.com";
const SHEET_ID = "12q1LcE-_LncMfeLraVr6YHjyI58-rhSkUKRneEls-WI";

let getHomepage = async (req, res) => {
  return res.render("homepage.ejs");
};

let getGoogleSheet = async (req, res) => {
  try {
    let currentDate = new Date();

    const format = "HH:mm DD/MM/YYYY";

    let formatedDate = moment(currentDate).format(format);

    // Initialize the sheet - doc ID is the long id in the sheets URL
    const doc = new GoogleSpreadsheet(SHEET_ID);

    // Initialize Auth - see more available options at https://theoephraim.github.io/node-google-spreadsheet/#/getting-started/authentication
    await doc.useServiceAccountAuth({
      client_email: CLIENT_EMAIL,
      private_key: PRIVATE_KEY,
    });

    await doc.loadInfo(); // loads document properties and worksheets

    const sheet = doc.sheetsByIndex[0]; // or use doc.sheetsById[id] or doc.sheetsByTitle[title]

    // append rows
    await sheet.addRow({
      "Faceboook nickname": "Tien Quang Dinh",
      Email: "tienquangdinh@gmail.com",
      "Phone number": `'0321456789`,
      "Time booked": formatedDate,
      "Customer's name": "Tien",
    });

    return res.send("Writing data to Google Sheet succeeds!");
  } catch (e) {
    return res.send(
      "Oops! Something wrongs, check logs console for detail ... "
    );
  }
};

module.exports = {
  getHomepage: getHomepage,
  getGoogleSheet: getGoogleSheet,
};
