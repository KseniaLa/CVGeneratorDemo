<template>
  <div class="home">
    <div class="columns">
      <div class="column">
        <div class="field">
          <div class="control">
            <label class="label">CV data</label>
            <input class="input is-info" v-model="firstName" type="text" placeholder="First Name" />
          </div>
        </div>
        <div class="field">
          <div class="control">
            <input class="input is-info" v-model="lastName" type="text" placeholder="Last Name" />
          </div>
        </div>
        <div class="field">
          <div class="control">
            <input class="input is-info" v-model="email" type="text" placeholder="Email" />
          </div>
        </div>
        <div class="field">
          <div class="control">
            <datepicker
              placeholder="Date of Birth"
              v-model="dateOfBirth"
              input-class="date-input"
              :clear-button="false"
            ></datepicker>
          </div>
        </div>
        <div class="field">
          <div class="control">
            <input class="input is-info" v-model="education" type="text" placeholder="Education" />
          </div>
        </div>
        <div class="field">
          <div class="control">
            <textarea
              class="textarea is-info"
              v-model="workingExperience"
              placeholder="Working experience"
            ></textarea>
          </div>
        </div>
        <button class="button is-info" v-on:click="generateCv">Generate</button>
      </div>
      <div class="column">
        <div class="field">
          <div class="control">
            <label class="label">CV style</label>
            <div class="select is-info">
              <select v-model="cvStyle">
                <option v-bind:value="0" hidden>Select style</option>
                <option
                  v-for="option in styleOptions"
                  v-bind:value="option.id"
                  v-bind:key="option.id"
                >{{ option.name }}</option>
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
import Datepicker from "vuejs-datepicker";
import ApiClient from "../api/apiClient.js";

export default {
  name: "Home",
  components: {
    Datepicker
  },
  data() {
    return {
      cvStyle: 0,
      styleOptions: [],
      firstName: "",
      lastName: "",
      email: "",
      dateOfBirth: null,
      education: "",
      workingExperience: ""
    };
  },
  methods: {
    getStyles: async function() {
      let styles = await ApiClient.getStyles();
      this.styleOptions = styles.data;
    },
    generateCv: async function() {
      let cvInfo = {
        styleId: this.cvStyle,
        cvInfo: {
          firstName: this.firstName,
          lastName: this.lastName,
          email: this.email,
          dateOfBirth: this.dateOfBirth,
          education: this.education,
          workingExperience: this.workingExperience
        }
      };
      await ApiClient.generateCv(cvInfo);
    }
  },
  mounted: async function() {
    await this.getStyles();
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