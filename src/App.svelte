<script>
  window['dbg'] = window.location.href.includes('?debug')
  window['dbg'] && console.log('\n---debug messaging enabled---\n\n')
  import Content from './Content.svelte'
  import Spinner from './components/Spinner.svelte'
  import { getAuth, accessToken } from './stores/auth/auth'
</script>

{#await getAuth()}
  <Spinner message={'Logging in...'} />
{:then _}
  <Content />
{/await}
