# New Project

✨ Project template using:

- [Svelte](https://svelte.dev)
- [Snowpack](https://snowpack.dev/)
- [Simple.css](https://simplecss.org/)
- [svelte-spa-router](https://github.com/ItalyPaleAle/svelte-spa-router)
- [Microsoft MSAL.js](https://docs.microsoft.com/en-us/azure/active-directory/develop/msal-overview)
- [Microsoft Graph API](https://docs.microsoft.com/en-us/azure/active-directory/develop/microsoft-graph-intro)
- [Office 365](https://www.office.com/)

This project is based on [laughing-barnacle](https://github.com/royashbrook/laughing-barnacle).

It provides an example of:
- A SPA application that
- Auto authenticates using MSAL.js upon startup
- Automatically gets a token for Microsoft Graph API upon successful authenication
- Provides a button to pull a value from a O365 site list by id and display it

Since CORS is not supported for SPA apps directly to a Azure Key Vault, this is showing how you could store that data in a sharepoint list and pull it.

While lack of CORS support with the Key Vault was the inspiration for creating this, everything here would work the same way if you wanted to pull list items from O365 for any other reason.

# Install

See [laughing-barnacle](https://github.com/royashbrook/laughing-barnacle) as the major steps are the same with a couple of notable differences:

1. The app registration must have permissions of `Sites.Read.All`
2. You need a sharepoint list

Below, I'll include details on how I added the site and the list.

## Create sharepoint site to hold the secret list

![image](https://user-images.githubusercontent.com/7390156/165871590-23a1c736-ce09-487a-9198-974157338097.png)

_note: this is a private site with only one list on it to hold the 'secret' data. this could also be a secured list on an existing site, but i figured i would include the process i used to test this._

## Create a new list to hold the secret data

![image](https://user-images.githubusercontent.com/7390156/165871660-aa7051fb-2dc7-47ee-99e4-6c5b76b2dd10.png)

## Add a secret item to the secret list

![image](https://user-images.githubusercontent.com/7390156/165958782-19ec8a45-37a6-43e1-9560-c0e17b217968.png)

We are going to pull it by the record ID in this example, so we don't need another field to put the data in. Sinced this is the first record, the ID will be 1. You could modify the graph query to pull multiple items, or search by name or use other fields, etc etc. But for the sake of this example this will do the trick.

## Get the path to use for the query using Microsoft Graph Explorer

I want to simply query this one record, so off to graph explorer to find all of my ids again. Let's search for my site:

![image](https://user-images.githubusercontent.com/7390156/165873121-517a08f9-1a57-42a0-abef-52771a2f8520.png)

So our site id is appears to be `ashbrookio.sharepoint.com,5d8c920c-cd55-4c5e-be46-784c50e0dded,591d0d2c-7744-45e9-ab45-053a32a8df87`

We can view lists on that site with:

https://graph.microsoft.com/v1.0/sites/ashbrookio.sharepoint.com,5d8c920c-cd55-4c5e-be46-784c50e0dded,591d0d2c-7744-45e9-ab45-053a32a8df87/lists

We can enumerate the items on that list with:

https://graph.microsoft.com/v1.0/sites/ashbrookio.sharepoint.com,5d8c920c-cd55-4c5e-be46-784c50e0dded,591d0d2c-7744-45e9-ab45-053a32a8df87/lists/secrets/items

We can reduce the data returned by adding some query parameters like:

https://graph.microsoft.com/v1.0/sites/ashbrookio.sharepoint.com,5d8c920c-cd55-4c5e-be46-784c50e0dded,591d0d2c-7744-45e9-ab45-053a32a8df87/lists/secrets/items/1?expand=fields(select=title)&$select=id

And, finally, we can remove all of the odata metadata by adding this header:

`headers.append("Accept", "application/json;odata.metadata=none");`

This will leave you with a nice clean'ish response like:

```
{
    "id": "1",
    "fields": {
        "Title": "I'm secret with ID1!"
    }
}
```

Unfortunately, there doesn't appear to be a way (that I could find) to *just* get one of the fields values in the response. I don't really need the ID, and would be perfectly happy with just getting the text value back for the property I want, but this does the trick. If anyone sees this and knows how to just get the text, please reach out. Everything would be the same, but I imagine we'd just update the query parameters above.

In my case, I just pulled the `Title` property out to display, and it looked like this after hitting the button:

![image](https://user-images.githubusercontent.com/7390156/165961585-6f1c0bb0-03e6-4cbd-993c-1f3360377480.png)
