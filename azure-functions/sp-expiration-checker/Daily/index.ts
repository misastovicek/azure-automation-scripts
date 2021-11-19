import { Context } from "@azure/functions"
import msal = require('@azure/msal-node');
import { AuthProvider, AuthProviderCallback, Client, Options } from "@microsoft/microsoft-graph-client"
import "isomorphic-fetch"
import axios from "axios";

interface IExpiringApplication {
    displayName: string
    id: string
    keyId: string
    keyType: string
    daysToExpire: number
    endDateTime: string
}

interface IAadApplication {
    id: string
    displayName: string
    keyCredentials: ICredential[]
    passwordCredentials: ICredential[]
}

interface ICredential {
    customKeyIdentifier: string
    displayName: string
    endDateTime: string
    startDateTime: string
    keyId: string
    hint?: string
    key?: string
    secretText?: string | null
    type?: string
    usage?: string
}

interface ITeamsBody {
    "@type": string
    "@context": string
    sections: ITeamsBodySection[]
    summary: string
    themeColor: string
}

interface ITeamsBodySection {
    activityTitle: string
    facts: ITeamsBodySectionFact[]
}

interface ITeamsBodySectionFact {
    name: string
    value: string
}

async function timerTrigger(context: Context, myTimer: any): Promise<void> {
    if (myTimer.isPastDue) {
        context.log('Timer function is running late!');
    }

    // Create App in App registration (Azure AD) and grant Application.Read.All API permissions to it (Application, not Delegated)
    // Then Grant consent to it so that the app can actually work
    const clientConfiguration: msal.Configuration = {
        auth: {
            clientId: process.env.CLIENT_ID,
            authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
            clientSecret: process.env.CLIENT_SECRET,
        }
    };

    const clientApp = new msal.ConfidentialClientApplication(clientConfiguration);
    const clientCredentialsRequest: msal.ClientCredentialRequest = {
        scopes: ["https://graph.microsoft.com/.default"]
    };


    const authProvider: AuthProvider = (callback: AuthProviderCallback) => {
        clientApp.acquireTokenByClientCredential(clientCredentialsRequest)
            .then((response) => {
                return callback(null, response.accessToken);
            })
            .catch((error) => {
                return callback(error, null);
            });
    };

    const options: Options = {
        authProvider
    };
    const client = Client.init(options);

    // Call to MS Graph
    const expirations = await getAppsKeysExpirations(await getAadApplications(client));
    if (expirations) {
        sendExpirationsToTeams(expirations);
    }
}

async function getTeamsBody(expiratingApplication: IExpiringApplication): Promise<ITeamsBody> {
    return {
        "@type": "MessageCard",
        "@context": "http://schema.org/extensions",
        summary: "Service Principal Expiration Warning!",
        themeColor: "0078D7",
        sections: [
            {
                activityTitle: `Service Principal Exires in ${expiratingApplication.daysToExpire} days!`,
                facts: [
                    {
                        name: "Application Name",
                        value: expiratingApplication.displayName
                    },
                    {
                        name: "Application ID",
                        value: expiratingApplication.id
                    },
                    {
                        name: "Key Type",
                        value: expiratingApplication.keyType
                    },
                    {
                        name: "Key ID",
                        value: expiratingApplication.keyId
                    },
                    {
                        name: "Expires at",
                        value: expiratingApplication.endDateTime
                    }
                ]
            }
        ]
    }
}

async function sendExpirationsToTeams(expiratingApplications: IExpiringApplication[]) {
    const teamsWebhook = process.env.TEAMS_WEBHOOK

    for (const app of expiratingApplications) {
        const teamsBody = await getTeamsBody(app)
        axios.post(teamsWebhook, teamsBody, {
            headers: {
                "content-type": "application/json",
            }
        }).then(data => console.log(`${data.status} - ${data.statusText}`))
            .catch(error => console.error(`Something went wrong: ${error}`))
    }
}

function getAadApplications(client: Client): Promise<IAadApplication[]> {
    return new Promise((resolve, reject) => {
        client.api("/applications?$select=id,displayName,keyCredentials,passwordCredentials").get()
            .then((response) => {
                return resolve(response.value);
            })
            .catch((error) => {
                console.error('\x1b[31m%s\x1b[0m', error["message"],);
                return reject(Array());
            });
    })
}

async function roundDateToDay(date: Date): Promise<Date> {
    return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), 0, 0, 0, 0);
}

async function getDaysToExpire(date: Date | null): Promise<number> {
    if (!date) {
        return -1;
    }

    const currentDate = await roundDateToDay(new Date())
    const expirationDays = [0, 1, 2, 3, 5, 7, 10, 14, 20, 25, 30]
    date = await roundDateToDay(date);

    for (const d of expirationDays) {
        const expirationDayInMs = d * 24 * 60 * 60 * 1000;
        if (currentDate.getTime() + expirationDayInMs === date.getTime()) {
            return d;
        }
    }
    return -1;
}

async function getAppsKeysExpirations(aadApps: IAadApplication[]): Promise<IExpiringApplication[]> {
    const aadAppsWithSecrets = aadApps.filter(x => {
        if (x.keyCredentials != null && x.keyCredentials.length > 0) {
            return x
        }
        if (x.passwordCredentials != null && x.passwordCredentials.length > 0) {
            return x
        }
    })

    let expiringApplications: IExpiringApplication[] = []
    for (const app of aadAppsWithSecrets) {
        expiringApplications = expiringApplications.concat(await getExpiringAppSecrets(app));
    }
    return expiringApplications
}

async function getExpiringAppSecrets(app: IAadApplication): Promise<IExpiringApplication[]> {
    const expiringApplications: IExpiringApplication[] = []

    for (const cred of app.keyCredentials.concat(app.passwordCredentials)) {
        const daysToExpire = await getDaysToExpire(new Date(cred.endDateTime));
        if (daysToExpire >= 0) {
            expiringApplications.push({
                id: app.id,
                displayName: app.displayName,
                keyId: cred.keyId,
                keyType: cred.type ?? "Password",
                daysToExpire: daysToExpire,
                endDateTime: cred.endDateTime
            });
        }
    }

    return expiringApplications
}

export default timerTrigger;