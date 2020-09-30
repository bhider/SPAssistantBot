Issue: Could not load Microsoft.Azure.WebJobs, version 2.2.0.0 
Solution: Upgrade Microsoft.NET.Sdk.Functions from 1.0.24 to 1.0.31

Issue: Could not load Microsoft.IdentityModel.Clients.ActiveDirectory 3.14.2
Solution: Added code to redirect assembly which is called at application startup (see ApplicationHelper.cs)