/* ===========================================================================
   BETA Otak — live dashboards configuration  (Google Drive source)
   ---------------------------------------------------------------------------
   Fill in the two things below, then host the "cloud" folder (e.g. drag it onto
   https://app.netlify.com/drop). The dashboards read the Excel files straight
   from your "Eng. Marwa" Drive folder and refresh by themselves when they change.
   =========================================================================== */
window.APP_CONFIG = {

  // 1) Your Google API key (Drive API enabled, restricted to your dashboard site).
  apiKey: "AIzaSyBxULKMrDiriwt_9Km7VNQyM_z6fr6ly04",

  // 2) The 3 Excel files in the "Eng. Marwa" folder, by their Drive file IDs.
  //    (The file ID is the long code in the file's share link.)
  sources: [
    { id: "1JUeDDl_6mzSwzpl7oDi22a3XRA-Y42Dp",  type: "subcontract" },
    { id: "1Li5qlA0DkZ_cBro2dOrCvTe1GfMlz2Gp",  type: "purchase"    },
    { id: "1vYpGYqd5BmNo2kYOR-e_s_LqCIGwWZWB",  type: "equipment"   }
  ],

  // How often (seconds) to check Drive for changes.
  pollSeconds: 10
};
