//https://docs.microsoft.com/en-us/graph/api/resources/driveitem?view=graph-rest-1.0
const URI = 'https://graph.microsoft.com/v1.0/drives'
const drive = 'b!UYlwtP7ji0euGnyoEVSB7CwNHVlEd-lFq0UFOjKo34eZk2vhVjpJTqRPC91hMyIA'
const item = '01XMKSJ4SHY645HMG6HJGLAI3BNS2V3PRK'
const sel = '$select=name,lastModifiedDateTime,content.downloadUrl'
export const graphContentEndpoint =
    `${URI}/${drive}/items/${item}/children?${sel}`