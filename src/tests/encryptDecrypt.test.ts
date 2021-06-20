/**
 * @jest-environment jsdom
 */

import "./setupTests";

import { smimeDecrypt, smimeEncrypt } from "../helpers/emailFunctions";
import { decodeHtml, encodeHtml } from "../helpers/converters";

// Import the same cert and key used within the add-in
import { cert, key } from "./certAndKey";

const plaintext = "This is some plaintext.";

test("encrypt and decrypt some plaintext", async () => {
  const encryptedText = await smimeEncrypt(plaintext, cert);
  const decryptedText = await smimeDecrypt(encryptedText, key, cert);
  expect(decryptedText).toBe(plaintext);
});

// The below tests exist since we surround S/MIME text with <pre> tags to prevent Outlook from inserting its own HTML tags within the S/MIME message and messing with the message integrity when setting the message body.
test("encrypt and decrypt some plaintext with <pre> tags surrounding", async () => {
  const encryptedText = await smimeEncrypt(plaintext, cert);
  const encryptedTextBody = "<pre>" + encryptedText + "</pre>";

  // Escape HTML encoded strings
  let originalEmailBody = decodeHtml(encryptedTextBody);

  // Remove <div>'s
  originalEmailBody = originalEmailBody.replace(/<div>/g, "");

  // Remove </div>'s
  originalEmailBody = originalEmailBody.replace(/<\/div>/g, "");

  // Remove <span>'s
  originalEmailBody = originalEmailBody.replace(/<span>/g, "");

  // Remove </span>'s
  originalEmailBody = originalEmailBody.replace(/<\/span>/g, "");

  // Replace <br>'s with \r\n
  originalEmailBody = originalEmailBody.replace(/<br>/g, "\r\n");

  // Remove <pre>'s
  originalEmailBody = originalEmailBody.replace(/<pre>/g, "");

  // Remove </pre>'s
  originalEmailBody = originalEmailBody.replace(/<\/pre>/g, "");

  // Detect S/MIME section
  let smimeSection = originalEmailBody.substring(originalEmailBody.indexOf("Content-Type:"));

  const decryptedText = await smimeDecrypt(smimeSection, key, cert);
  expect(decryptedText).toBe(plaintext);
});

test("encrypt and decrypt text surrounded by <span> and <pre> tags", async () => {
  const encryptedText = await smimeEncrypt(plaintext, cert);
  const encryptedTextPre = "<pre>" + encryptedText + "</pre>";
  const encryptedTextSpanPre = "<pre>" + encryptedTextPre + "</span>";

  // Escape HTML encoded strings
  let originalEmailBody = decodeHtml(encryptedTextSpanPre);

  // Remove <div>'s
  originalEmailBody = originalEmailBody.replace(/<div>/g, "");

  // Remove </div>'s
  originalEmailBody = originalEmailBody.replace(/<\/div>/g, "");

  // Remove <span>'s
  originalEmailBody = originalEmailBody.replace(/<span>/g, "");

  // Remove </span>'s
  originalEmailBody = originalEmailBody.replace(/<\/span>/g, "");

  // Replace <br>'s with \r\n
  originalEmailBody = originalEmailBody.replace(/<br>/g, "\r\n");

  // Remove <pre>'s
  originalEmailBody = originalEmailBody.replace(/<pre>/g, "");

  // Remove </pre>'s
  originalEmailBody = originalEmailBody.replace(/<\/pre>/g, "");

  // Detect S/MIME section
  let smimeSection = originalEmailBody.substring(originalEmailBody.indexOf("Content-Type:"));

  const decryptedText = await smimeDecrypt(smimeSection, key, cert);
  expect(decryptedText).toBe(plaintext);
});

// Outlook HTML encodes
test("encrypt and decrypt text surrounded by <span> and <pre> tags that has been HTML encoded", async () => {
  const encryptedText = await smimeEncrypt(plaintext, cert);
  // console.log(encryptedText);
  const encryptedTextHtmlEncoded = encodeHtml(encryptedText);
  // console.log(encryptedTextHtmlEncoded);
  const encryptedTextPre = "<pre>" + encryptedTextHtmlEncoded + "</pre>";
  // console.log(encryptedTextPre);
  const encryptedTextSpanPre = "<span>" + encryptedTextPre + "</span>";
  // console.log(encryptedTextSpanPre);

  // Escape HTML encoded strings
  let originalEmailBody = decodeHtml(encryptedTextSpanPre);
  // let originalEmailBody = decodeHtml(encryptedTextSpanPre);

  // Remove <div>'s
  originalEmailBody = originalEmailBody.replace(/<div>/g, "");

  // Remove </div>'s
  originalEmailBody = originalEmailBody.replace(/<\/div>/g, "");

  // Remove <span>'s
  originalEmailBody = originalEmailBody.replace(/<span>/g, "");

  // Remove </span>'s
  originalEmailBody = originalEmailBody.replace(/<\/span>/g, "");

  // Replace <br>'s with \r\n
  originalEmailBody = originalEmailBody.replace(/<br>/g, "\r\n");

  // Remove <pre>'s
  originalEmailBody = originalEmailBody.replace(/<pre>/g, "");

  // Remove </pre>'s
  originalEmailBody = originalEmailBody.replace(/<\/pre>/g, "");

  // Detect S/MIME section
  let smimeSection = originalEmailBody.substring(originalEmailBody.indexOf("Content-Type:"));

  const decryptedText = await smimeDecrypt(smimeSection, key, cert);
  // console.log(decryptedText)
  expect(decryptedText).toBe(plaintext);
});
