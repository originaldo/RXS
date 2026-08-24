# Opt-in Methods (CTA) — paste into the GoTo "Opt-in methods" box

Customers opt in to receive text messages from America's Restoration Contractor (ARC)
in one of two ways:

1. **Website contact form (arcsvcs.com):** When requesting service, the customer enters
   their mobile number and checks a dedicated SMS-consent checkbox that is unchecked by
   default and separate from any email or phone-call consent. The checkbox reads:
   "I agree to receive service-related text messages from America's Restoration Contractor
   (ARC) at the number provided. Message frequency varies. Msg & data rates may apply.
   Reply STOP to opt out, HELP for help. See our Privacy Policy and Terms."

2. **Inbound service request via a lead platform (e.g., Thumbtack, Yelp):** The customer
   initiates contact and provides their mobile number to request restoration services;
   ARC responds by text to coordinate the requested job. Consent is limited to
   service-related messages about the customer's own request.

Messages are transactional/customer-care only (appointment scheduling and reminders,
technician arrival notifications, job and insurance-claim status updates, and replies to
customer questions). No marketing or promotional messages are sent under this campaign.
Opt-in data and consent are never shared with third parties.

Privacy Policy: https://arcsvcs.com/privacy
Terms & Conditions: https://arcsvcs.com/terms

---

## Website SMS-consent checkbox — HTML to add to the arcsvcs.com contact form

<label>
  <input type="checkbox" name="sms_consent" value="yes">
  I agree to receive service-related text messages from America's Restoration Contractor (ARC)
  at the number provided. Message frequency varies. Msg &amp; data rates may apply.
  Reply STOP to opt out, HELP for help. See our
  <a href="/privacy">Privacy Policy</a> and <a href="/terms">Terms</a>.
</label>

Notes:
- Keep it UNCHECKED by default.
- Keep it SEPARATE from any email/phone consent checkbox (SMS-only).
- Store the timestamp + the number + which checkbox was ticked (proof of consent).
