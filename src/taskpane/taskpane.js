/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // document.getElementById("run").onclick = run;
    document.getElementById("placeholder_feature_button").onclick = () => {
      window.location.href = "https://techdevalex.github.io/Excel_Addin_Hosting/src/taskpane/feature_placeholder/feature_placeholder.html";

    }
  }
});
