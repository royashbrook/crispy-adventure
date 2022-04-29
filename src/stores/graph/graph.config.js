//https://docs.microsoft.com/en-us/graph/api/resources/listitem?view=graph-rest-1.0
const URI = 'https://graph.microsoft.com/v1.0/sites'
const site = 'ashbrookio.sharepoint.com,5d8c920c-cd55-4c5e-be46-784c50e0dded,591d0d2c-7744-45e9-ab45-053a32a8df87'
const list = 'secrets'
const item = '1'
const qry = 'expand=fields(select=title)&$select=id'
export const graphContentEndpoint =
    `${URI}/${site}/lists/${list}/items/${item}?${qry}`