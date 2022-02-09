"use strict";

(function() {
  Office.onReady(function() {
    // Office is ready
    $(document).ready(function() {
      // The document is ready
      checkSignature();
    });
  });

  function checkSignature() {
    alert("checkSignature triggered");
    var item = Office.context.mailbox.item;
    var user_profile = Office.context.mailbox.userProfile;
    var letterHeadTemplate = getTemplate();
    letterHeadTemplate = letterHeadTemplate.replace("{Full_Name}", user_profile.displayName);
    Office.context.mailbox.item.body.setAsync(letterHeadTemplate, { coercionType: Office.CoercionType.Html }, function(
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
})();

$("#set-selected-data").click(setSelectedData);

function setSelectedData() {}

function getTemplate() {
  //var template = "<b> Hi Team,</b><b><br>How is everyone?</b><br><br><b>Ready for Outlook web add-in demo? </b>";
  var template =
    '<html xmlns="http://www.w3.org/1999/xhtml"><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"><title>State Farm eLetterhead</title><style type="text/css">body,td,th{font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#0000}a{color:#686869;text-decoration:none;font-weight:700}a:hover{color:#cc0717}</style></head><body><div id="letter_template"><table width="650" border="0" cellspacing="0" cellpadding="0"><tr height="55" bgcolor="#cc0717"><td colspan="7"><img src="Images/eLet-BANNER.gif" alt="State Farm Agent" width="650" height="55"></td></tr><tr><td width="212" align="left" valign="top" bgcolor="#ededed"><img src="Images/eLet-PHOTO.gif" alt="Photo" width="212" height="210"><br><table width="212" cellspacing="0" cellpadding="4" bgcolor="#ededed"><tr><td colspan="2" style="text-align:center;padding-bottom:16px" width="100%"><span style="color:#cc0717;font-weight:700;font-size:13px">{Full_Name}</span><br><span style="font-weight:700;color:#686869;font-size:12px">{Job_Type}</span><br></td></tr><tr><td style="border-top:solid 1px #e0e0e0" width="20%"><img src="https://portal.brandmyemail.com/bme-images/icons-email" alt="Email Me" width="36" height="30"></td><td width="80%" style="border-top:solid 1px #e0e0e0;font-weight:700"><a href="Mailto:{Email_Address}">Email Me!</a></td></tr><tr><td style="border-top:solid 1px #e0e0e0" width="20%"><img src="https://portal.brandmyemail.com/bme-images/icons-phone" alt="Phone" width="36" height="30"></td><td width="80%" style="border-top:solid 1px #e0e0e0;font-weight:700"><a href="tel://{Telephone_Number}">{Telephone_Number}</a></td></tr><tr><td style="border-top:solid 1px #e0e0e0" width="20%"><img src="https://portal.brandmyemail.com/bme-images/icons-mobile" alt="Text Me" width="38" height="26"></td><td width="80%" style="border-top:solid 1px #e0e0e0;font-weight:700"><a href="sms://{Mobile_Number}">Text Me!</a></td></tr><tr><td style="border-top:solid 1px #e0e0e0" width="20%"><img src="https://portal.brandmyemail.com/bme-images/icons-fax" alt="Fax" width="38" height="26"></td><td width="80%" style="border-top:solid 1px #e0e0e0;font-weight:700"><a href="tel://{Fax_Number}">{Fax_Number}</a></td></tr><tr><td style="border-top:solid 1px #e0e0e0" width="20%"><img src="https://portal.brandmyemail.com/bme-images/icons-agent" alt="Home Page" width="36" height="30"></td><td width="80%" style="border-top:solid 1px #e0e0e0"><a href="{Website_Url}" target="_blank">Visit Joe Home Page</a></td></tr><tr><td style="border-top:solid 1px #e0e0e0" width="20%"><img src="https://portal.brandmyemail.com/bme-images/icons-directions" alt="Map and Directions" width="36" height="30"></td><td width="80%" style="border-top:solid 1px #e0e0e0"><a href="#">Map &amp; Directions</a></td></tr><tr><td style="" width="20 % "><img src="https://portal.brandmyemail.com/bme-images/icons-account" alt="Account" width="36" height="30"></td><td width="80%" style="border-top:solid 1px #e0e0e0"><a href="https://oams.statefarm.com/auth/UI/Login/">Access Your Account</a></td></tr><tr><td style="border-top:solid 1px #e0e0e0;border-bottom:solid 1px #e0e0e0" width="20%"><img src="https://portal.brandmyemail.com/bme-images/icons-quote" alt="Request A Quote" width="38" height="26"></td><td width="80%" style="border-top:solid 1px #e0e0e0;border-bottom:solid 1px #e0e0e0"><a href="https://www.statefarm.com/agent/xxxxxxxxxxx/quote">Request a Quote</a></td></tr><tr><td colspan="2" align="center"><p><span style="color:#686869;font-size:12px">987 N. America St.<br>Anywhere, WI 54952-4321</span><br><br><span style="color:#686869;font-weight:700;font-size:13px">Languages</span><br><br><span style="color:#cc0717;font-weight:700;font-size:13px">NMLS #</span></p></td></tr></table></td><td colspan="6" valign="top"><table width="435" border="0" cellpadding="0" cellspacing="0"><tr><td style="padding:20px"><span style="color:#000;font-family:Calibri,sans-serif;font-size:14.5px"><p>Hello Agent,</p><p>This is a sample letter to clients! This is just to show you how your email will look once you type in a message to your client or email recipient.<br><br><strong>NOTE - The font used here is Calibri 11pt, which is our standard font we use. If you would like a special font and size, you must let us know. Should you get your final files and need to change fonts at a later time, you will be charged an editing fee to change your files.</strong></p><p>We hope you like the layout!</p><p>Please thoroughly check through ALL of your contact information and make sure we did not miss anything. Also, check ALL links, to make sure they are working properly.</p><p>That is what this proofing phase is for: to make sure we catch any issues before we finalize your eStationery files.</p><p>Thank you!</p></span><p><span style="color:#cc0717;font-family:Arial,Helvetica,sans-serif;font-weight:700;font-size:13px">Joseph J Agent</span><br><span style="color:#686869;font-family:Arial,Helvetica,sans-serif;font-size:12px">Agent</span></p><p style="color:#cc0717;font-family:Arial,Helvetica,sans-serif;font-weight:700;font-size:13px"><a href="http://www.website.com">www.MyWebsite.com</a><br><br><a href="https://www.statefarm.com/insurance/life"><img src="https://portal.brandmyemail.com/bme-images/life-insurance-calculator" width="60" height="67" alt="Life Insurance Calculator"></a>&nbsp; &nbsp; &nbsp;<a href="https://www.statefarm.com/simple-insights/auto-and-vehicles/calculators"><img src="https://portal.brandmyemail.com/bme-images/car-loan-calculator" width="60" height="67" alt="Car Loan Calculator"></a>&nbsp; &nbsp; &nbsp;<a href="https://www.statefarm.com/simple-insights/retirement/calculators"><img src="https://portal.brandmyemail.com/bme-images/retirement-calculator" width="60" height="67" alt="Retirement Calculator"></a>&nbsp; &nbsp; &nbsp;<br><br><img src="https://portal.brandmyemail.com/bme-images/google-reviewus-color" width="60" height="67" alt="Review Us on Google"> &nbsp; &nbsp; &nbsp; <img src="https://portal.brandmyemail.com/bme-images/yelp-findus" width="60" height="67" alt="Find Us on Yelp"> &nbsp; &nbsp; &nbsp; <img src="https://portal.brandmyemail.com/bme-images/facebook-likeus" width="60" height="67" alt="Like Us on Facebook"> &nbsp; &nbsp; &nbsp;<br><br><img src="https://portal.brandmyemail.com/bme-images/join-our-team" width="60" height="67" alt="Join Our Team"> &nbsp; &nbsp; &nbsp;<a href="https://www.statefarm.com/customer-care/download-mobile-apps/state-farm-mobile-app"><img src="https://portal.brandmyemail.com/bme-images/mobile-app" width="60" height="67" alt="Mobile App"></a>&nbsp; &nbsp; &nbsp;<a href="https://www.statefarm.com/finances/banking"><img src="https://portal.brandmyemail.com/bme-images/banking" width="60" height="67" alt="Online &amp; Mobile Banking"></a>&nbsp; &nbsp; &nbsp;<br><br><a href="https://www.statefarm.com/finances/banking/loans/home-loans"><img src="https://portal.brandmyemail.com/bme-images/button-mortgage" width="60" height="67" alt="Mortgage"></a>&nbsp; &nbsp; &nbsp; <img src="https://portal.brandmyemail.com/bme-images/button-budget-calculator" width="60" height="67" alt="Budget Calculator"> &nbsp; &nbsp; &nbsp;<a href="https://financials.statefarm.com/UnauthenticatedPaymentsUI-web/unauth/launchApp.do"><img src="https://portal.brandmyemail.com/bme-images/button-pay-a-bill" width="60" height="67" alt="Pay a Bill"></a>&nbsp; &nbsp; &nbsp;<br><br><a href="https://www.statefarm.com/finances/banking/credit-cards#sf-rewards"><img src="https://portal.brandmyemail.com/bme-images/button-credit-card" width="60" height="67" alt="State Farm Credit Card"></a>&nbsp; &nbsp; &nbsp;<a href="https://www.statefarm.com/customer-care/download-mobile-apps/drive-safe-and-save-mobile"><img src="https://portal.brandmyemail.com/bme-images/button-drive-safe-and-save" width="60" height="67" alt="Drive Safe &amp; Save"></a>&nbsp; &nbsp; &nbsp;<a href="https://www.statefarm.com/insurance/auto/discounts"><img src="https://portal.brandmyemail.com/bme-images/button-auto-discounts" alt="Auto Discounts" width="60" height="67"></a><br><br></p></td></tr></table></td></tr><tr><td colspan="7"><img border="0" width="650" height="38" src="https://portal.brandmyemail.com/bme-images/sf-eLet-Footer" alt="Providing Insurance and Financial Services" v:shapes="Picture_x0020_7"></td></tr><tr><td width="212"><img src="https://portal.brandmyemail.com/bme-images/sm-bar-gradient" alt="Gray bar" width="212" height="34" border="0" v:shapes="Picture_x0020_7"></td><td width="264"><img src="https://portal.brandmyemail.com/bme-images/sm-bar-stay-connected-264" alt="Stay Connected" width="264" height="34" border="0" v:shapes="Picture_x0020_7"></td><td width="35"><a href="https://www.facebook.com/statefarm" target="_blank"><img src="https://portal.brandmyemail.com/bme-images/sm-bar-facebook" alt="Facebook" width="35" height="34" border="0" v:shapes="Picture_x0020_7"></a></td><td width="36"><a href="https://twitter.com/statefarm" target="_blank"><img src="https://portal.brandmyemail.com/bme-images/sm-bar-twitter" alt="Twitter" width="36" height="34" border="0" v:shapes="Picture_x0020_7"></a></td><td width="35"><a href="https://www.youtube.com/statefarm" target="_blank"><img src="https://portal.brandmyemail.com/bme-images/sm-bar-youtube2" alt="YouTube" width="35" height="34" border="0" v:shapes="Picture_x0020_7"></a></td><td width="33"><a href="https://www.linkedin.com/company/2381/" target="_blank"><img src="https://portal.brandmyemail.com/bme-images/sm-bar-linkedin" alt="LinkedIn" width="33" height="34" border="0" v:shapes="Picture_x0020_7"></a></td><td width="35" height="34"><a href="https://www.instagram.com/statefarm/" target="_blank"><img border="0" width="35" height="34" src="https://portal.brandmyemail.com/bme-images/sm-bar-instagram" alt="Instagram" v:shapes="Picture_x0020_7"></a></td></tr><tr><td colspan="7" height="16"><a href="https://www.brandmyemail.com/"><img border="0" src="https://portal.brandmyemail.com/bme-images/powered-by" width="220" height="16" alt="Powered By Brand My Email"></a></td></tr></table></div></body></html>';

  return template;
}
