<template>
	<header class="app-header level is-marginless">
		<div class="app-name level-left">
			<router-link to="/">CV Generator</router-link>
		</div>
	</header>
</template>

<script>
export default {
  name: "AppHeader",
};
</script>

<style lang="scss">

.app-header {
  padding: 10px;
  background-color: #3298dc;
  color: #ffffff;
	box-shadow: 0 4px 8px 0 #00000033, 0 6px 20px 0 #00000030;
}

.app-name {
  color: #ffffff;
  a {
    font-weight: bold;
    color: #ffffff;
		font-size: 50px;
		text-decoration: none;

		&:hover	{
			color: #ffffff;
		}
  }
}
</style>