/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { getUserProfile, sendMailAsUser } from "../helpers/sso-helper";
import { filterUserProfileInfo } from "./../helpers/documentHelper";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("forward").onclick = onForwardButtonClicked;
  }
});

export async function onForwardButtonClicked() {
  const item = Office.context.mailbox.item;
  item.body.getAsync(Office.CoercionType.Html, function callback(result) {
    sendMailAsUser(handleSendMailResponse, "providencesatterfield@gmail.com", item.subject, {
      content: result.value,
      contentType: "HTML",
    });
  });
}

function handleSendMailResponse(response) {
  // console.log(response);
}

// function writeDataToOfficeDocument(result) {
//   let data = [];
//   let userProfileInfo = filterUserProfileInfo(result);

//   for (let i = 0; i < userProfileInfo.length; i++) {
//     if (userProfileInfo[i] !== null) {
//       data.push(userProfileInfo[i]);
//     }
//   }

//   let userInfo = "";
//   for (let i = 0; i < data.length; i++) {
//     userInfo += data[i] + "\n";
//   }

//   Office.context.mailbox.item.body.setSelectedDataAsync(userInfo, { coercionType: Office.CoercionType.Html });
// }
