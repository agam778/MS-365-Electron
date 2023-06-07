window.addEventListener("DOMContentLoaded", () => {
  const releaseTagElement = document.getElementById("release-tag");

  fetch("https://api.github.com/repos/agam778/MS-365-Electron/releases/latest")
    .then((response) => response.json())
    .then((data) => {
      const latestTag = data.tag_name;
      releaseTagElement.textContent = `${latestTag}`;
    })
    .catch((error) => {
      console.error("Error:", error);
      releaseTagElement.textContent = "Failed to fetch release tag";
    });
});
