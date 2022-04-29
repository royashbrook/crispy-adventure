export async function callMSGraph(endpoint, token) {

    const headers = new Headers();
    const bearer = `Bearer ${token}`;

    headers.append("Authorization", bearer);
    headers.append("Accept", "application/json;odata.metadata=none");

    const options = {
        method: "GET",
        headers: headers
    };

    try {
        const response = await fetch(endpoint, options);
        const data = await response.json();
        return data;//.value;
    } catch (error) {
        return console.error(error);
    }
}