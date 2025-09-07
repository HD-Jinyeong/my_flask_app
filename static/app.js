function addFileInput() {
  const div = document.getElementById("fileInputs");
  const input = document.createElement("input");
  input.type = "file";
  input.name = "files";
  div.appendChild(document.createElement("br"));
  div.appendChild(input);
}
