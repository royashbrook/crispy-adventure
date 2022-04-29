import { writable,get } from 'svelte/store';
export const data = writable(null);

import { accessToken } from '../auth/auth';
import { graphContentEndpoint } from "./graph.config";
import { callMSGraph } from './graph';

export const getData = async () => {
  
  let graphdata = await callMSGraph(graphContentEndpoint, get(accessToken));
  data.set(graphdata.fields.Title);

}
