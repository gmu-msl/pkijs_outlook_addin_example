import { Crypto } from "@peculiar/webcrypto";
import * as asn1js from "asn1js";
import { Convert } from "pvtsutils";
import * as pkijs from "pkijs";

// Set crypto engine
const crypto = new Crypto();
const engineName = "@peculiar/webcrypto";
pkijs.setEngine(
  engineName,
  crypto,
  new pkijs.CryptoEngine({ name: engineName, crypto: crypto, subtle: crypto.subtle })
);

import { PemConverter } from "./converters";

import MimeNode from "emailjs-mime-builder";
import smimeParse from "emailjs-mime-parser";

/**
 * Adapted from PKI.js' SMIMEEncryptionExample
 * @returns {string} encrypted string
 * @param {string} text string to encrypt
 * @param {string} certificatePem public certificate to encrypt with in PEM format
 * @param {string} oaepHashAlgo algorithm to hash the text with (defaults to SHA-256)
 * @param {string} encryptionAlgo algorithm to encrypt the text with ("AES-CBC" or "AES-GCM")
 * @param {Number} length length to encrypt the text to (default 128)
 */
export async function smimeEncrypt(
  text: string,
  certificatePem: string,
  oaepHashAlgo: string = "SHA-256",
  encryptionAlgo: string = "AES-CBC",
  length: Number = 128
): Promise<string> {
  // Decode input certificate
  const asn1 = asn1js.fromBER(PemConverter.decode(certificatePem)[0]);
  const certSimpl = new pkijs.Certificate({ schema: asn1.result });

  const cmsEnveloped = new pkijs.EnvelopedData();

  cmsEnveloped.addRecipientByCertificate(certSimpl, { oaepHashAlgorithm: oaepHashAlgo });

  await cmsEnveloped.encrypt({ name: encryptionAlgo, length: length }, Convert.FromUtf8String(text));

  const cmsContentSimpl = new pkijs.ContentInfo();
  cmsContentSimpl.contentType = "1.2.840.113549.1.7.3";
  cmsContentSimpl.content = cmsEnveloped.toSchema();

  const schema = cmsContentSimpl.toSchema();
  const ber = schema.toBER(false);

  // Insert enveloped data into new Mime message
  const mimeBuilder = new MimeNode("application/pkcs7-mime; name=smime.p7m; smime-type=enveloped-data; charset=binary")
    .setHeader("content-description", "Enveloped Data")
    .setHeader("content-disposition", "attachment; filename=smime.p7m")
    .setHeader("content-transfer-encoding", "base64")
    .setContent(new Uint8Array(ber));
  mimeBuilder.setHeader("from", "sender@example.com");
  mimeBuilder.setHeader("to", "recipient@example.com");
  mimeBuilder.setHeader("subject", "Example S/MIME encrypted message");

  return mimeBuilder.build();
}
/**
 * Adapted from PKI.js' SMIMEEncryptionExample
 * @returns {string} decrypted string
 * @param {string} text string to decrypt
 * @param {string} privateKeyPem user's private key to decrypt with in PEM format
 * @param {string} certificatePem user's public certificate to decrypt with in PEM format
 */
export async function smimeDecrypt(text: string, privateKeyPem: string, certificatePem: string): Promise<string> {
  // Decode input certificate
  let asn1 = asn1js.fromBER(PemConverter.decode(certificatePem)[0]);
  const certSimpl = new pkijs.Certificate({ schema: asn1.result });

  // Decode input private key
  const privateKeyBuffer = PemConverter.decode(privateKeyPem)[0];

  // Parse S/MIME message to get CMS enveloped content
  try {
    const parser = smimeParse(text);

    // Make all CMS data
    asn1 = asn1js.fromBER(parser.content.buffer);
    if (asn1.offset === -1) {
      alert('Unable to parse your data. Please check you have "Content-Type: charset=binary" in your S/MIME message');
      return;
    }

    const cmsContentSimpl = new pkijs.ContentInfo({ schema: asn1.result });
    const cmsEnvelopedSimpl = new pkijs.EnvelopedData({ schema: cmsContentSimpl.content });

    const message = await cmsEnvelopedSimpl.decrypt(0, {
      recipientCertificate: certSimpl,
      recipientPrivateKey: privateKeyBuffer,
    });

    return Convert.ToUtf8String(message);
  } catch (err) {
    // Not an S/MIME message
    throw err;
  }
}
