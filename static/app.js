const form = document.querySelector("#certificate-form");
const statusMessage = document.querySelector("#status");
const button = document.querySelector("#submit-button");

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const formData = new FormData(form);
  const name = formData.get("name").trim();

  button.disabled = true;
  button.textContent = "Generating…";
  statusMessage.textContent = "Creating your editable certificate…";
  statusMessage.className = "status";

  try {
    const response = await fetch("/api/generate", { method: "POST", body: formData });
    if (!response.ok) {
      const error = await response.json().catch(() => ({}));
      throw new Error(error.detail || "Something went wrong. Please try again.");
    }

    const file = await response.blob();
    const url = URL.createObjectURL(file);
    const link = document.createElement("a");
    const safeName = name.replace(/[<>:"/\\|?*]+/g, "").trim().replace(/\s+/g, "_") || "certificate";
    link.href = url;
    link.download = `${safeName}_Certificate.pptx`;
    link.click();
    URL.revokeObjectURL(url);

    const replacements = Number(response.headers.get("X-GenerateX-Replacements"));
    statusMessage.textContent = replacements
      ? "Done — your editable certificate has downloaded."
      : "Your file has downloaded, but {{NAME}} was not found. Check the template placeholder.";
    statusMessage.className = replacements ? "status success" : "status warning";
  } catch (error) {
    statusMessage.textContent = error.message;
    statusMessage.className = "status error";
  } finally {
    button.disabled = false;
    button.innerHTML = "Generate editable PPTX <span>→</span>";
  }
});
