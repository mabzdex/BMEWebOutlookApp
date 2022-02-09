// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Contains code for event-based activation on both Outlook on web and Outlook on Windows.

/**
 * Checks if signature exists.
 * If not, displays the information bar for user.
 * If exists, insert signature into new item (appointment or message).
 * @param {*} eventObj Office event object
 * @returns
 */
function checkSignature(eventObj) {
  var item = Office.context.mailbox.item;
  Office.context.mailbox.item.body.setAsync(getTemplate(), { coercionType: Office.CoercionType.Html }, function(
    asyncResult
  ) {
    console.log(asyncResult.status);
    console.log(Office.AsyncResultStatus.Succeeded);
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Selected text has been updated successfully.");
    } else {
      console.error(asyncResult.error);
    }
  });
}

function getTemplate() {
  //var template = "<b> Hi Team,</b><b><br>How is everyone?</b><br><br><b>Ready for Outlook web add-in demo? </b>";
  var template =
    '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8" /><title>State Farm eLetterhead</title><style type="text/css"><!--body,td,th {font-family: Arial, Helvetica, sans-serif;font-size: 12px;color: #0000;}a {color:#686869; text-decoration:none;font-weight:bold; }a:hover{color:#cc0717;}--></style></head><body><table width="650" border="0" cellspacing="0" cellpadding="0"> <tr height="55" bgcolor="#cc0717"> <td colspan="7"><img src="Images/eLet-BANNER.gif" alt="State Farm Agent" width="650" height="55" /></td> </tr> <tr> <td width="212" align="left" valign="top" bgcolor="#ededed"><img src="Images/eLet-PHOTO.gif" alt="Photo" width="212" height="210" /><br/> <table width="212" cellspacing="0" cellpadding="4" bgcolor="#ededed"> <tr> <td colspan="2" style="text-align: center; padding-bottom:16px;" width="100%"><span style="color: #cc0717; font-weight: bold; font-size: 13px;">Joseph J Agent</span><br /> <span style="font-weight: bold; color: #686869; font-size: 12px;">Agent</span><br /></td> </tr> <tr> <td style="border-top: solid 1px #e0e0e0;" width="20%"><img src="https://portal.brandmyemail.com/bme-images/icons-email" alt="Email Me" width="36" height="30" /></td> <td width="80%" style="border-top: solid 1px #e0e0e0; font-weight: bold;"><a href="Mailto:XXXXXXXX@statefarm.com">Email Me!</a></td> </tr> <tr> <td style="border-top: solid 1px #e0e0e0;" width="20%"><img src="https://portal.brandmyemail.com/bme-images/icons-phone" alt="Phone" width="36" height="30" /></td> <td width="80%" style="border-top: solid 1px #e0e0e0; font-weight: bold;"><a href="tel://555-555-5555">555-555-5555</a></td> </tr> <tr> <td style="border-top: solid 1px #e0e0e0;" width="20%"><img src="https://portal.brandmyemail.com/bme-images/icons-mobile" alt="Text Me" width="38" height="26" /></td> <td width="80%" style="border-top: solid 1px #e0e0e0; font-weight: bold;"><a href="sms://555-555-5555">Text Me!</a></td> </tr> <tr> <td style="border-top: solid 1px #e0e0e0;" width="20%"><img src="https://portal.brandmyemail.com/bme-images/icons-fax" alt="Fax" width="38" height="26" /></td> <td width="80%" style="border-top: solid 1px #e0e0e0; font-weight: bold;"><a href="tel://555-555-5555">555-555-5555</a></td> </tr> <tr> <td style="border-top: solid 1px #e0e0e0;" width="20%"><img src="https://portal.brandmyemail.com/bme-images/icons-agent" alt="Home Page" width="36" height="30" /></td> <td width="80%" style="border-top: solid 1px #e0e0e0;"><a href="http://www.statefarm.com/" target="_blank">Visit Joe Home Page</a></td > </tr> <tr> <td style="border-top: solid 1px #e0e0e0;" width="20%"><img src="https:/ / portal.brandmyemail.com / bme - images / icons - directions" alt="Map and Directions" width="36" height="30" /></td> <td width="80 % " style="border - top: solid 1px #e0e0e0; "><a href="#">Map &amp; Directions</a></td> </tr> <tr> <td style="border - top: solid 1px #e0e0e0; " width="20 % "><img src="https://portal.brandmyemail.com/bme-images/icons-account" alt="Account" width="36" height="30" /></td> <td width="80%" style="border-top: solid 1px #e0e0e0; "><a href="https://oams.statefarm.com/auth/UI/Login/">Access Your Account</a></td> </tr> <tr> <td style="border-top: solid 1px #e0e0e0; border-bottom: solid 1px #e0e0e0;" width="20%"><img src="https://portal.brandmyemail.com/bme-images/icons-quote" alt="Request A Quote" width="38" height="26" /></td> <td width="80%" style="border-top: solid 1px #e0e0e0; border-bottom: solid 1px #e0e0e0;"><a href="https://www.statefarm.com/agent/xxxxxxxxxxx/quote">Request a Quote</a></td> </tr> <tr> <td colspan="2" align="center"> <p><span style="color: #686869; font-size: 12px;"> 987 N. America St.<br/> Anywhere, WI 54952-4321</span><br/><br/> <span style="color: #686869; font-weight: bold; font-size: 13px;">Languages</span><br/><br/> <span style="color: #cc0717; font-weight: bold; font-size: 13px;">NMLS #</span></p></td> </tr> </table> </td> <td colspan="6" valign="top"> <table width="435" border="0" cellpadding="0" cellspacing="0"> <tr> <td style="padding:20px;"> <span style="color: #000000; font-family: Calibri, sans-serif; font-size: 14.5px;"> <p>Hello Agent,</p> <p>This is a sample letter to clients! This is just to show you how your email will look once you type in a message to your client or email recipient.<br/> <br/><strong>NOTE - The font used here is Calibri 11pt, which is our standard font we use. If you would like a special font and size, you must let us know. Should you get your final files and need to change fonts at a later time, you will be charged an editing fee to change your files.</strong></p> <p>We hope you like the layout! </p> <p>Please thoroughly check through ALL of your contact information and make sure we did not miss anything. Also, check ALL links, to make sure they are working properly. </p> <p>That is what this proofing phase is for: to make sure we catch any issues before we finalize your eStationery files. </p> <p>Thank you!</p></span> <p><span style="color: #cc0717; font-family:Arial, Helvetica, sans-serif;font-weight: bold; font-size: 13px;">Joseph J Agent</span><br> <span style="color: #686869;font-family:Arial, Helvetica, sans-serif; font-size: 12px;">Agent</span></p> <p style="color: #cc0717;font-family:Arial, Helvetica, sans-serif; font-weight: bold; font-size: 13px;"><a href="http://www.website.com">www.MyWebsite.com</a><br/><br/> <a href="https://www.statefarm.com/insurance/life"><img src="https://portal.brandmyemail.com/bme-images/life-insurance-calculator" width="60" height="67" alt="Life Insurance Calculator"/></a> &nbsp; &nbsp; &nbsp; <a href="https://www.statefarm.com/simple-insights/auto-and-vehicles/calculators"><img src="https://portal.brandmyemail.com/bme-images/car-loan-calculator" width="60" height="67" alt="Car Loan Calculator"/></a> &nbsp; &nbsp; &nbsp; <a href="https://www.statefarm.com/simple-insights/retirement/calculators"><img src="https://portal.brandmyemail.com/bme-images/retirement-calculator" width="60" height="67" alt="Retirement Calculator"/></a> &nbsp; &nbsp; &nbsp; <br/><br/> <img src="https://portal.brandmyemail.com/bme-images/google-reviewus-color" width="60" height="67" alt="Review Us on Google"/> &nbsp; &nbsp; &nbsp; <img src="https://portal.brandmyemail.com/bme-images/yelp-findus" width="60" height="67" alt="Find Us on Yelp"/> &nbsp; &nbsp; &nbsp; <img src="https://portal.brandmyemail.com/bme-images/facebook-likeus" width="60" height="67" alt="Like Us on Facebook"/> &nbsp; &nbsp; &nbsp; <br/><br/> <img src="https://portal.brandmyemail.com/bme-images/join-our-team" width="60" height="67" alt="Join Our Team"/> &nbsp; &nbsp; &nbsp; <a href="https://www.statefarm.com/customer-care/download-mobile-apps/state-farm-mobile-app"><img src="https://portal.brandmyemail.com/bme-images/mobile-app" width="60" height="67" alt="Mobile App"/></a> &nbsp; &nbsp; &nbsp; <a href="https://www.statefarm.com/finances/banking"><img src="https://portal.brandmyemail.com/bme-images/banking" width="60" height="67" alt="Online &amp; Mobile Banking"/></a> &nbsp; &nbsp; &nbsp; <br/><br/> <a href="https://www.statefarm.com/finances/banking/loans/home-loans"><img src="https://portal.brandmyemail.com/bme-images/button-mortgage" width="60" height="67" alt="Mortgage"/></a> &nbsp; &nbsp; &nbsp; <img src="https://portal.brandmyemail.com/bme-images/button-budget-calculator" width="60" height="67" alt="Budget Calculator"/> &nbsp; &nbsp; &nbsp; <a href="https://financials.statefarm.com/UnauthenticatedPaymentsUI-web/unauth/launchApp.do"><img src="https://portal.brandmyemail.com/bme-images/button-pay-a-bill" width="60" height="67" alt="Pay a Bill"/></a> &nbsp; &nbsp; &nbsp; <br/><br/> <a href="https://www.statefarm.com/finances/banking/credit-cards#sf-rewards"><img src="https://portal.brandmyemail.com/bme-images/button-credit-card" width="60" height="67" alt="State Farm Credit Card"/></a> &nbsp; &nbsp; &nbsp; <a href="https://www.statefarm.com/customer-care/download-mobile-apps/drive-safe-and-save-mobile"><img src="https://portal.brandmyemail.com/bme-images/button-drive-safe-and-save" width="60" height="67" alt="Drive Safe &amp; Save"/></a> &nbsp; &nbsp; &nbsp; <a href="https://www.statefarm.com/insurance/auto/discounts"><img src="https://portal.brandmyemail.com/bme-images/button-auto-discounts" alt="Auto Discounts" width="60" height="67"/></a> <br/><br/> </p></td> </tr></table> </td> </tr> <tr> <td colspan="7"><img border="0" width="650" height="38" src="https://portal.brandmyemail.com/bme-images/sf-eLet-Footer" alt="Providing Insurance and Financial Services"v:shapes="Picture_x0020_7" /></td> </tr> <tr> <td width="212"><img src="https://portal.brandmyemail.com/bme-images/sm-bar-gradient" alt="Gray bar" width="212" height="34" border="0"v:shapes="Picture_x0020_7" /></td> <td width="264"><img src="https://portal.brandmyemail.com/bme-images/sm-bar-stay-connected-264" alt="Stay Connected" width="264" height="34" border="0"v:shapes="Picture_x0020_7" /></td> <td width="35"><a href="https://www.facebook.com/statefarm" target="_blank"><img src="https://portal.brandmyemail.com/bme-images/sm-bar-facebook" alt="Facebook" width="35" height="34" border="0"v:shapes="Picture_x0020_7" /></a></td> <td width="36"><a href="https://twitter.com/statefarm" target="_blank"><img src="https://portal.brandmyemail.com/bme-images/sm-bar-twitter" alt="Twitter" width="36" height="34" border="0"v:shapes="Picture_x0020_7" /></a></td> <td width="35"><a href="https://www.youtube.com/statefarm" target="_blank"><img src="https://portal.brandmyemail.com/bme-images/sm-bar-youtube2" alt="YouTube" width="35" height="34" border="0"v:shapes="Picture_x0020_7" /></a></td> <td width="33"><a href="https://www.linkedin.com/company/2381/" target="_blank"><img src="https://portal.brandmyemail.com/bme-images/sm-bar-linkedin" alt="LinkedIn" width="33" height="34" border="0"v:shapes="Picture_x0020_7" /></a></td> <td width="35" height="34"><a href="https://www.instagram.com/statefarm/" target="_blank"><img border="0" width="35" height="34" src="https://portal.brandmyemail.com/bme-images/sm-bar-instagram" alt="Instagram"v:shapes="Picture_x0020_7" /></a></td> </tr> <tr> <td colspan="7" height="16"><a href="https://www.brandmyemail.com/"><img border=0 src="https://portal.brandmyemail.com/bme-images/powered-by" width="220" height="16" alt="Powered By Brand My Email"></a></td> </tr></table></body></html>';

  return template;
}

/**
 * For Outlook on Windows only. Insert signature into appointment or message.
 * Outlook on Windows can use setSignatureAsync method on appointments and messages.
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @param {*} user_info Information details about the user
 * @param {*} eventObj Office event object
 */
function insert_auto_signature(compose_type, user_info, eventObj) {
  let template_name = get_template_name(compose_type);
  let signature_info = get_signature_info(template_name, user_info);
  addTemplateSignature(signature_info, eventObj);
}

/**
 * 
 * @param {*} signatureDetails object containing:
 *  "signature": The signature HTML of the template,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 * @param {*} eventObj 
 * @param {*} signatureImageBase64 
 */
function addTemplateSignature(signatureDetails, eventObj, signatureImageBase64) {
  if (is_valid_data(signatureDetails.logoBase64) === true) {
    //If a base64 image was passed we need to attach it.
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
      signatureDetails.logoBase64,
      signatureDetails.logoFileName,
      {
        isInline: true,
      },
      function (result) {
        //After image is attached, insert the signature
        Office.context.mailbox.item.body.setSignatureAsync(
          signatureDetails.signature,
          {
            coercionType: "html",
            asyncContext: eventObj,
          },
          function (asyncResult) {
            asyncResult.asyncContext.completed();
          }
        );
      }
    );
  } else {
    //Image is not embedded, or is referenced from template HTML
    Office.context.mailbox.item.body.setSignatureAsync(
      signatureDetails.signature,
      {
        coercionType: "html",
        asyncContext: eventObj,
      },
      function (asyncResult) {
        asyncResult.asyncContext.completed();
      }
    );
  }
}

/**
 * Creates information bar to display when new message or appointment is created
 */
function display_insight_infobar() {
  Office.context.mailbox.item.notificationMessages.addAsync("fd90eb33431b46f58a68720c36154b4a", {
    type: "insightMessage",
    message: "Please set your signature with the PnP sample add-in.",
    icon: "Icon.16x16",
    actions: [
      {
        actionType: "showTaskPane",
        actionText: "Set signatures",
        commandId: get_command_id(),
        contextData: "{''}",
      },
    ],
  });
}

/**
 * Gets template name (A,B,C) mapped based on the compose type
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @returns Name of the template to use for the compose type
 */
function get_template_name(compose_type) {
  if (compose_type === "reply") return Office.context.roamingSettings.get("reply");
  if (compose_type === "forward") return Office.context.roamingSettings.get("forward");
  return Office.context.roamingSettings.get("newMail");
}

/**
 * Gets HTML signature in requested template format for given user
 * @param {\} template_name Which template format to use (A,B,C)
 * @param {*} user_info Information details about the user
 * @returns HTML signature in requested template format
 */
function get_signature_info(template_name, user_info) {
  if (template_name === "templateB") return get_template_B_info(user_info);
  if (template_name === "templateC") return get_template_C_info(user_info);
  return get_template_A_info(user_info);
}

/**
 * Gets correct command id to match to item type (appointment or message)
 * @returns The command id
 */
function get_command_id() {
  if (Office.context.mailbox.item.itemType == "appointment") {
    return "MRCS_TpBtn1";
  }
  return "MRCS_TpBtn0";
}

/**
 * Gets HTML string for template A
 * Embeds the signature logo image into the HTML string
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template A,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 */
function get_template_A_info(user_info) {
  const logoFileName = "sample-logo.png";
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str += "<tr>";
  // Embed the logo using <img src='cid:...
  str +=
    "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='cid:" +
    logoFileName +
    "' alt='MS Logo' width='24' height='24' /></td>";
  str += "<td style='padding-left: 5px;'>";
  str += "<strong>" + user_info.name + "</strong>";
  str += is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str += "<br/>";
  str += is_valid_data(user_info.job) ? user_info.job + "<br/>" : "";
  str += user_info.email + "<br/>";
  str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str += "</td>";
  str += "</tr>";
  str += "</table>";

  // return object with signature HTML, logo image base64 string, and filename to reference it with.
  return {
    signature: str,
    logoBase64:
      "iVBORw0KGgoAAAANSUhEUgAAACIAAAAiCAYAAAA6RwvCAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAEeSURBVFhHzdhBEoIwDIVh4EoeQJd6YrceQM+kvo5hQNokLymO/4aF0/ajlBl1fL4bEp0uj3K9XQ/lGi0MEcB3UdD0uVK1EEj7TIuGeBaKYCgIswCLcUMid8mMcUEiCMk71oRYE+Etsd4UD0aFeBBSFtOEMAgpg6lCIggpitlAMggpgllBeiAkFjNDeiIkBlMgeyAkL6Z6WJdlEJJnjvF4vje/BvRALNN23tyRXzVpd22dHSZtLhjMHemB8cxRINZZyGCssbL2vCN7YLwItHo0PTEMAm3OSA8Mi0DVw5rBRBCoCkERTBSBmhDEYDII5PqlZy1iZSGQuiOSZ6JW3rEuCIpgmDFuCGImZuEUBHkWiOweDUHaQhEE+pM/aobhBZaOpYLJeeeoAAAAAElFTkSuQmCC",
    logoFileName: logoFileName,
  };
}

/**
 * Gets HTML string for template B
 * References the signature logo image from the HTML
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template B,
    "logoBase64": null since this template references the image and does not embed it ,
    "logoFileName": null since this template references the image and does not embed it
 */
function get_template_B_info(user_info) {
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str += "<tr>";
  // Reference the logo using a URI to the web server <img src='https://...
  str +=
    "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://officedev.github.io/Office-Add-in-samples/Samples/outlook-set-signature/assets/sample-logo.png' alt='Logo' /></td>";
  str += "<td style='padding-left: 5px;'>";
  str += "<strong>" + user_info.name + "</strong>";
  str += is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str += "<br/>";
  str += user_info.email + "<br/>";
  str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str += "</td>";
  str += "</tr>";
  str += "</table>";

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Gets HTML string for template C
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template C,
    "logoBase64": null since there is no image,
    "logoFileName": null since there is no image
 */
function get_template_C_info(user_info) {
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += user_info.name;

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Validates if str parameter contains text.
 * @param {*} str String to validate
 * @returns true if string is valid; otherwise, false.
 */
function is_valid_data(str) {
  return str !== null && str !== undefined && str !== "";
}

Office.actions.associate("checkSignature", checkSignature);
