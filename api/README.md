
# Azure Function: contact (Microsoft Graph SendMail)

This function accepts POSTs from your website contact form and sends an email from your mailbox (**MAILBOX_ADDRESS**) via Microsoft Graph.

## Deploy (quick)
1. Create a **Function App** (Node 18) in Azure (Consumption plan is fine).
2. Deploy: GitHub Actions via Deployment Center, or zip deploy.
3. App Settings → **Configuration**:
   - `TENANT_ID` = your Entra tenant GUID
   - `CLIENT_ID` = App registration (Application) ID
   - `CLIENT_SECRET` = client secret value (keep safe)
   - `MAILBOX_ADDRESS` = e.g., `info@schneiderdrafting.com`

## App Registration (Entra ID)
1. New app registration → single-tenant.
2. Certificates & secrets → New client secret.
3. API permissions → **Microsoft Graph** → **Application permissions** → **Mail.Send**.
4. Grant admin consent.

## Website hookup
- Your contact form posts to `/api/contact` (already updated in **index.cal.azure-contact.html**).
- Function is **anonymous** to allow browser posts. Consider adding WAF/Front Door, basic rate-limit, or CAPTCHA if needed.
