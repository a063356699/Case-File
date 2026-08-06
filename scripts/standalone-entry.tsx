import React from "react";
import { createRoot } from "react-dom/client";
import Home from "../app/page";

document.title = "\u7e3d\u8868 \u55ae\u6a5fhtml";
if (localStorage.getItem("property-desk-v1") === null) {
  localStorage.setItem("property-desk-v1", "[]");
}
createRoot(document.getElementById("root")!).render(<Home />);
