# Microsoft Teams Group Calendar Tab sample

The following sample is a Microsoft Teams tab application that will show a Group Calendar by calling the Microsoft Graph API  to get the members of a particular Team, then the getSchedule Graph operation to get the schedule of the users involved. The Display side of the application makes use of the [FullCalendar](https://fullcalendar.io/) JavaScript event calendar. The legend is built using the users photo which is downloaded from the Graph API also.

Screen Shots of the Tab application in Action 

Month View

![](https://gscales.github.io/TeamsGroupCalendar/docs/gcScreen1.JPG)

Week View

![](https://gscales.github.io/TeamsGroupCalendar/docs/gcscren3.JPG)

Day View 

![](https://gscales.github.io/TeamsGroupCalendar/docs/gcscren2.JPG)

List View

![](https://gscales.github.io/TeamsGroupCalendar/docs/gcscren4.JPG)

# User Permission requirements #

As this app runs as the currently logged on user that user must have a least the Calendar Details (Free/Busy time, subject, location) Freebusy permission to be able to view calendar events from Team members mailboxes.

# **Installation** #

**Prerequisites for using Teams Tab Applications**

To use a Teams Tab application application side loading must be enabled in the Office365 portal see the following page for how to modify the Teams Org setting [https://docs.microsoft.com/en-us/microsoftteams/admin-settings](https://docs.microsoft.com/en-us/microsoftteams/admin-settings). 
> "Sideloading is how you add an app to Teams by uploading a zip file directly to a team. Side-loading lets you test an app as it's being developed. It also lets you build an app for internal use only and share it with your team without submitting it to the Teams app catalog in the Office Store. "

![](https://gscales.github.io/TeamsGroupCalendar/docs/Sideloading.JPG)

**Note**: Make sure you use the https://admin.microsoft.com/AdminPortal/Home#/Settings/ServicesAndAddIns and not the Teams Admin portal as you won't be able to finding this setting in the later.

# Testing this GitHub Instance #

The application files for a Teams Tab application needs to be hosted on a web server, for testing only you can use this hosted version on gitHub. To use this you would need to grant the following applicationId consent in your tenant using the following URL

[https://login.microsoftonline.com/common/adminconsent?client_id=71db5de8-6b7d-437c-b973-0e13f81619e8](https://login.microsoftonline.com/common/adminconsent?client_id=71db5de8-6b7d-437c-b973-0e13f81619e8)

You then need to download the Manifest Zip file from [https://github.com/gscales/gscales.github.io/raw/master/TeamsGroupCalendar/TabPackage/app.zip
](https://github.com/gscales/gscales.github.io/raw/master/TeamsGroupCalendar/TabPackage/app.zip)
then follow the Custom App installation process described below


# **Custom App Installation Process** #

Official documentation for installing Custom Apps can be found 
[https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/apps/apps-upload](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/apps/apps-upload)

Walk through

Choose the Microsoft Store Icon in the Team client (don't worry you not going to purchase anything)
![](https://gscales.github.io/TeamsGroupCalendar/docs/walkthrough1.JPG)

Then select "Upload Custom Application" & "For Me and My Teams"

![](https://gscales.github.io/TeamsGroupCalendar/docs/walkthrough2.JPG)

Select the Teams you want to install the app into 

![](https://gscales.github.io/TeamsGroupCalendar/docs/walkthrough3.JPG)

Then select the Channel to install the Tab onto.

# Hosting the Application yourself #

Create an Application Registration that has the following grants

![](https://gscales.github.io/TeamsGroupCalendar/docs/grantsrequired.JPG)

Modify the manifest of the Application registration to enable the Implicit authentication flow 

    "logoutUrl": null,
  	"oauth2AllowImplicitFlow": true,
    "oauth2AllowUrlPathMatching": false,

Change the tab application m** Manifest** (your version of [https://github.com/gscales/gscales.github.io/blob/master/TeamsGroupCalendar/TabPackage/manifest.json](https://github.com/gscales/gscales.github.io/blob/master/TeamsGroupCalendar/TabPackage/manifest.json))

You need to change the Id,PackageName and configurationURL setting in the manifest to your own unique ApplicationId and URL where the config.html page is hosted

      "$schema": "https://statics.teams.microsoft.com/sdk/v1.2/manifest/MicrosoftTeams.schema.json", 
  	  "manifestVersion": "1.3",
      "version": "1.0.0",
      "id": "71db5de8-6b7d-437c-b973-0e13f81619e8",
      "packageName": "TeamsGroupCalendar.io.github.gscales",
    configurableTabs": [
    {
      "configurationUrl": "https://gscales.github.io/TeamsGroupCalendar/app/config.html",
      "canUpdateConfiguration": true,
      "scopes": [ "team" ]
    }
    ],

Modify you hosted version of the https://github.com/gscales/gscales.github.io/blob/master/TeamsGroupCalendar/app/Config/appconfig.js file. Change the clientId to the applicationId from your application registration and the hostRoot to the root of your webhost.

     const getConfig = () => {
  	 var config = {
        clientId : "71db5de8-6b7d-437c-b973-0e13f81619e8",
        redirectUri : "/TeamsGroupCalendar/app/silent-end.html",
        authwindow :  "/TeamsGroupCalendar/app/auth.html",
	 hostRoot: "https://gscales.github.io",
   	 };
  	 return config;
	}

Create a Zip file of all the files in the https://github.com/gscales/gscales.github.io/blob/master/TeamsGroupCalendar/TabPackage directory and then use that in the Custom App Installation process described above.













