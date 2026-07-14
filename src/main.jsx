import React from "react";
import ReactDOM from "react-dom/client";
import AttendanceLedger from "./AttendanceLedger.jsx";
import "./storage-polyfill.js";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <AttendanceLedger />
  </React.StrictMode>
);
