<script>
    import { accessToken, user } from '../stores/auth/auth'
    import { getData, data} from '../stores/graph/data';
</script>

<h2>Welcome, {$user}!</h2>
<h4>Congratulations! You got a token!</h4>
<p><code>{$accessToken?.substring(0,20) + '...'}</code></p>
<button on:click={getData}>Do something with the token!</button>
{#if $data}
<pre width="100%">{JSON.stringify($data,null,2)}</pre>
{/if}
