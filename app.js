Office.onReady(() => {
  const urlInput = document.getElementById("urlInput");
  const loadButton = document.getElementById("loadButton");
  const configButton = document.getElementById("configButton");
  const setup = document.getElementById("setup");
  const viewer = document.getElementById("viewer");
  const pollFrame = document.getElementById("pollFrame");

  const savedUrl = localStorage.getItem("pollUrl");

  if (savedUrl) {
    showViewer(savedUrl);
  }

  loadButton.addEventListener("click", () => {
    const url = urlInput.value.trim();

    if (!url.startsWith("https://")) {
      alert("Use uma URL iniciando com https://");
      return;
    }

    localStorage.setItem("pollUrl", url);
    showViewer(url);
  });

  configButton.addEventListener("click", () => {
    setup.classList.remove("hidden");
    viewer.classList.add("hidden");
    urlInput.value = localStorage.getItem("pollUrl") || "";
  });

  function showViewer(url) {
    pollFrame.src = url;
    setup.classList.add("hidden");
    viewer.classList.remove("hidden");
  }
});