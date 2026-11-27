const RELEASE = {
  version: "v1.13.0",
  textSize: "101.5 MB",
  ocrSize: "101.5 MB",
  textSha: "7BFA2C461CE9FBFC36195F91804A118209EA09FB030F5ADCA223995315D67632",
  ocrSha: "D87160B637078642C0013A028426A9A986B98BCAC00B11D082D0C6164281C297",
};

const root = document.documentElement;
const copyStatus = document.querySelector(".copy-status");

function applyLanguage(language) {
  const normalized = language === "en" ? "en" : "zh";
  root.dataset.uiLang = normalized;
  root.lang = normalized === "en" ? "en" : "zh-CN";
  document.querySelectorAll("[data-set-lang]").forEach((button) => {
    button.setAttribute(
      "aria-pressed",
      String(button.dataset.setLang === normalized),
    );
  });
  try {
    localStorage.setItem("kindle-release-language", normalized);
  } catch {
    // Language still applies for this session when storage is unavailable.
  }
}

function showCopyStatus(message) {
  if (!copyStatus) return;
  copyStatus.textContent = message;
  copyStatus.classList.add("visible");
  window.clearTimeout(showCopyStatus.timer);
  showCopyStatus.timer = window.setTimeout(() => {
    copyStatus.classList.remove("visible");
  }, 1800);
}

async function copyChecksum(key) {
  const checksum = key === "ocr" ? RELEASE.ocrSha : RELEASE.textSha;
  try {
    await navigator.clipboard.writeText(checksum);
    showCopyStatus(root.dataset.uiLang === "en" ? "Checksum copied" : "校验值已复制");
  } catch {
    showCopyStatus(root.dataset.uiLang === "en" ? "Copy failed" : "复制失败");
  }
}

document.querySelectorAll("[data-set-lang]").forEach((button) => {
  button.addEventListener("click", () => applyLanguage(button.dataset.setLang));
});

document.querySelectorAll("[data-release-size]").forEach((element) => {
  element.textContent =
    element.dataset.releaseSize === "ocr" ? RELEASE.ocrSize : RELEASE.textSize;
});

document.querySelectorAll("[data-release-sha]").forEach((element) => {
  element.textContent =
    element.dataset.releaseSha === "ocr" ? RELEASE.ocrSha : RELEASE.textSha;
});

document.querySelectorAll("[data-copy-sha]").forEach((button) => {
  button.addEventListener("click", () => copyChecksum(button.dataset.copySha));
});

let savedLanguage = "zh";
try {
  savedLanguage = localStorage.getItem("kindle-release-language") || "zh";
} catch {
  savedLanguage = "zh";
}
applyLanguage(savedLanguage);
