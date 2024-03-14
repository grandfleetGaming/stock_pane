/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
import fetch from "cross-fetch";
import { OPENAI_KEY, ALPHAVANTAGE_KEY } from "./env";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function openAiChat(prompt) {
  const apiKey = OPENAI_KEY;
  const questionPrompt = `
  "Act as an investment advisor. For the following input, please recommend a single stock ticker to research or buy. Return your response in the following format: 
            
  {
    "ticker": "..."
  }

  Make sure you only return json.
  `;

  const response = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: "gpt-3.5-turbo",
      messages: [
        {
          role: "system",
          content: questionPrompt,
        },
        { role: "user", content: prompt },
      ],
      temperature: 0.7,
    }),
  });

  const data = await response.json();
  return data.choices[0].message.content;
}

export async function alphaVantage(symbol) {
  var url = `https://www.alphavantage.co/query?function=OVERVIEW&symbol=${symbol}&apikey=${ALPHAVANTAGE_KEY}`;
  const response = await fetch(url);
  const data = await response.json();
  return data;
}

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const inputValue = document.getElementById("prompt").value;
      // const range = context.workbook.getSelectedRange();

      // Read the range address
      // range.load("address");

      // Update the fill color
      // range.format.fill.color = "yellow";
      // cell contents fill with inputValue
      await context.sync();
      const stockResponse = await openAiChat(inputValue);
      console.log(stockResponse);
      // console.log(`The range address was ${range.address}.`);
      console.log("Done., inputValue", inputValue);
      // set id="response" html element to response

      // parse as json
      const tickerData = JSON.parse(stockResponse);
      const results = await alphaVantage(tickerData.ticker);
      console.log(results);
      const description = results.Description;
      document.getElementById("description").innerHTML = description;
    });
  } catch (error) {
    console.error(error);
  }
}
