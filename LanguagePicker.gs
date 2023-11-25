<html>
  <head>
    <base target="_top">
  </head>
  <body>
    <div>
      <label for="language-select">Choose a language:</label>
      <select id="language-select" name="languages">
        <option value="German">German</option>
        <option value="Dutch">Dutch</option>
        <option value="French">French</option>
        <option value="Spanish">Spanish</option>
        <option value="Italian">Italian</option>
        <option value="Portuguese">Portuguese</option>
        <option value="Russian">Russian</option>
        <option value="Polish">Polish</option>
        <option value="Catalan">Catalan</option>
        <option value="Swedish">Swedish</option>
        <option value="Hungarian">Hungarian</option>
        <option value="Arabic">Arabic</option>
        <option value="Hebrew">Hebrew</option>
        <option value="Indonesian">Indonesian</option>
        <option value="Turkish">Turkish</option>
        <option value="Japanese">Japanese</option>
        <option value="English">English</option>
      </select>
    </div>
    <button onclick="setLanguage()">Set Language</button>

    <script>
  function setLanguage() {
    var language = document.getElementById('language-select').value;
    google.script.run.withSuccessHandler(languageSet)
                      .updateLanguage(language);
  }

  function languageSet() {
    // Optional: Add any actions to perform after the language is set
    google.script.host.close();
  }
</script>


  </body>
</html>
