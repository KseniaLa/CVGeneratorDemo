<template>
  <div class="home">
    <div class="columns">
      <div class="column">
        <div class="field">
          <div class="control">
            <label class="label">CV data</label>
            <input class="input is-info" type="text" placeholder="First Name" />
          </div>
        </div>
        <div class="field">
          <div class="control">
            <input class="input is-info" type="text" placeholder="Last Name" />
          </div>
        </div>
        <div class="field">
          <div class="control">
            <datepicker placeholder="Date of Birth" v-model="dueDate" input-class="date-input" :clear-button="false"></datepicker>
          </div>
        </div>
        <div class="field">
          <div class="control">
            <input class="input is-info" type="text" placeholder="Education" />
          </div>
        </div>
        <div class="field">
          <div class="control">
            <textarea class="textarea is-info" placeholder="Working experience"></textarea>
          </div>
        </div>
        <button class="button is-info">Generate</button>
      </div>
      <div class="column">
        <div class="field">
          <div class="control">
            <label class="label">CV style</label>
            <div class="select is-info">
              <select>
                <option>Select style</option>
                <option>With options</option>
              </select>
            </div>
          </div>
        </div>
      </div>
      <div class="column"></div>
    </div>
  </div>
</template>

<script>
import Datepicker from 'vuejs-datepicker';

export default {
  name: "Home",
  components: {
    Datepicker
  }
};
</script>

<style lang="scss">
.home {
  padding: 20px;
}

.date-input {
  width: 100%;
  padding: 5px;
  border-radius: 5px;
  border-width: 1px;
  border-color: #3298dc;
  text-decoration: none;
  outline: none;
  font-size: 15px;
}
</style>