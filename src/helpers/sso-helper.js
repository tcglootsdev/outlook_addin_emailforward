/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { dialogFallback } from "./fallbackauthdialog.js";
import { getUserData, sendMail } from "./middle-tier-calls";
import { showMessage } from "./message-helper";
import { handleClientSideErrors } from "./error-handler";

/* global Office */

let retryGetMiddletierToken = 0;

let recipient = null;
let subject = null;
let body = null;

export async function getUserProfile(callback) {
  try {
    let middletierToken = await Office.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    });
    let response = await getUserData(middletierToken);
    if (!response) {
      throw new Error("Middle tier didn't respond");
    } else if (response.claims) {
      // Microsoft Graph requires an additional form of authentication. Have the Office host
      // get a new token using the Claims string, which tells AAD to prompt the user for all
      // required forms of authentication.
      let mfaMiddletierToken = await Office.auth.getAccessToken({
        authChallenge: response.claims,
      });
      response = getUserData(mfaMiddletierToken);
    }

    // AAD errors are returned to the client with HTTP code 200, so they do not trigger
    // the catch block below.
    if (response.error) {
      handleAADErrors(response, callback);
    } else {
      callback(response);
    }
  } catch (exception) {
    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        dialogFallback(callback, recipient, subject, body);
      }
    } else {
      showMessage("EXCEPTION: " + JSON.stringify(exception));
      throw exception;
    }
  }
}

export async function sendMailAsUser(callback, ...args) {
  recipient = args[0];
  subject = args[1];
  body = args[2];
  try {
    let middletierToken = await Office.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    });
    let response = await sendMail(middletierToken, recipient, subject, body);
    if (!response) {
      throw new Error("Middle tier didn't respond");
    } else if (response.claims) {
      // Microsoft Graph requires an additional form of authentication. Have the Office host
      // get a new token using the Claims string, which tells AAD to prompt the user for all
      // required forms of authentication.
      let mfaMiddletierToken = await Office.auth.getAccessToken({
        authChallenge: response.claims,
      });
      response = sendMail(mfaMiddletierToken, recipient, subject, body);
    }

    // AAD errors are returned to the client with HTTP code 200, so they do not trigger
    // the catch block below.
    if (response.error) {
      handleAADErrors(response, callback);
    } else {
      callback(response);
    }
  } catch (exception) {
    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        dialogFallback(callback, recipient, subject, body);
      }
    } else {
      showMessage("EXCEPTION: " + JSON.stringify(exception));
      throw exception;
    }
  }
}

function handleAADErrors(response, callback) {
  // On rare occasions the middle tier token is unexpired when Office validates it,
  // but expires by the time it is sent to AAD for exchange. AAD will respond
  // with "The provided value for the 'assertion' is not valid. The assertion has expired."
  // Retry the call of getAccessToken (no more than once). This time Office will return a
  // new unexpired middle tier token.

  if (response.error_description.indexOf("AADSTS500133") !== -1 && retryGetMiddletierToken <= 0) {
    retryGetMiddletierToken++;
    // getUserProfile(callback);
    sendMail(callback, recipient, subject, body);
  } else {
    dialogFallback(callback, recipient, subject, body);
  }
}
