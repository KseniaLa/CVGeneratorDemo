import axios from "axios";
import { API_BASE, GET_STYLE, GENERATE_CV } from "./apiConstants.js";

class ApiClient {
  constructor() {}

  static async getStyles() {
    let result = null;

    let url = `${API_BASE}${GET_STYLE}`;
    console.log(url)
    try {
      result = await axios.get(url, {
        headers: {}
      });
    } catch {
      return { data: [], success: false };
    }

    return { data: result.data, success: true };
  }

  static async generateCv(cvInfo) {
    //let data = JSON.stringify(cvInfo);

    let url = `${API_BASE}${GENERATE_CV}`;

    axios({
      url: url,
      method: "POST",
      data: cvInfo,
      responseType: "blob"
    }).then(response => {
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement("a");
      link.href = url;
      link.setAttribute("download", "CV.docx");
      document.body.appendChild(link);
      link.click();
    });
  }
}

export default ApiClient;
