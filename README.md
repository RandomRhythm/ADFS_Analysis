## AD FS Research

This repo contains information about AD FS logging that I researched years ago. There are also some example scripts that were used to review the described events.

### Claims Activity Parsing

The 500 and 501 event IDs contain additional information around claims activity such as identify. The following descriptors capture identify information such as username.

- http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn
- http://schemas.xmlsoap.org/ws/2005/05/identity/claims/implicitupn
- http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name
- http://schemas.xmlsoap.org/claims/UPN

The 500 and 501 events also include an instance id, which correlates to other events.

**Instance ID:**

**1b033855-c665-4531-a710-28a32bd45f9b**

The instance ID can be used to correlate to event IDs 299, 324, and 412. The 299 ID documents a successfully issued token while 324 is a token issuance failure. The 299 and 324 event IDs also include an Instance ID. The activity ID, similar to the instance ID, links different events together. These events link instance ID and Activity ID together and track the failure/success of authentication. These events also document the party relying on the request.

**Instance ID: 1b033855-c665-4531-a710-28a32bd45f9b**

**Activity ID: 00000000-0000-0000-de03-0080000000a1**

**Relying party: urn:federation:MicrosoftOnline**

The event 412 documents a successful token authentication and contains both an instance ID and activity ID. The 411 ID documented failed token authentication and only provides an activity ID. However, the activity ID for both appeared to be all zeros for the observed events. The 411 event ID data does not correlate with the other events.

**Instance ID: f4a66e1a-d2d6-488c-a292-5ffd0d06ceae**

**Activity ID: 00000000-0000-0000-0000-000000000000**

The 410 and 413 IDs also have an Activity ID. Event ID 410 provides the request context headers associated with an Activity ID, which includes user agent, client application and forwarded client IP. The 413 event ID provides diagnostic information around token authentication error, which includes the activity ID, caller, relying party, client IP, etc.

### Failed Authentication

The 411 event ID did not correlate with the reported claims activity. Within the event were the email address, client IP address and the exception details. The exception details sometimes contains the reason for the failure but other times it did not provide useful information.

**Token validation failed. See inner exception for more details.**

**Additional Data**

**Activity ID: 00000000-0000-0000-0000-000000000000**

**Token Type:**

**http://schemas.microsoft.com/ws/2006/05/identitymodel/tokens/UserName**

**Client IP:**

**11.11.11.11,172.16.36.42**

**Error message:**

**user@domain.com**

**Exception details:**

**System.IdentityModel.Tokens.SecurityTokenValidationException: user@domain.com**

**at Microsoft.IdentityServer.Service.Tokens.MSISWindowsUserNameSecurityTokenHandler.ValidateToken(SecurityToken token)**

### Locked Accounts

AD FS logs to event ID 512 and 515 for locked account events. ID 512 documents a failed password login attempt to a locked account. It provides the email address, client IP, bad password count and the last bad password attempt. ID 515 documents a successful password was attempted against a locked account.

**The following user account was in a locked out state and the correct password was just provided. This account may be compromised.**

**Additional Data**

**Activity ID: 00000000-0000-0000-0000-000000000000**

**User:**

**user@domain.com**

**Device Certificate:**

**-**

**Client IP:**

**77.77.77.77,172.16.36.42**