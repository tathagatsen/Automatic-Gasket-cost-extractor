document.addEventListener("DOMContentLoaded", function () {
    function setupDropArea(dropAreaId, fileInputId) {
        const dropArea = document.getElementById(dropAreaId);
        const fileInput = document.getElementById(fileInputId);

        dropArea.addEventListener("click", () => fileInput.click());

        dropArea.addEventListener("dragover", (event) => {
            event.preventDefault();
            dropArea.style.backgroundColor = "#e0e0e0";
        });

        dropArea.addEventListener("dragleave", () => {
            dropArea.style.backgroundColor = "";
        });

        dropArea.addEventListener("drop", (event) => {
            event.preventDefault();
            dropArea.style.backgroundColor = "";
            const files = event.dataTransfer.files;
            if (files.length > 0) {
                dropArea.innerHTML = `<p>${files[0].name}</p>`;
            }
        });

        fileInput.addEventListener("change", (event) => {
            if (fileInput.files.length > 0) {
                dropArea.innerHTML = `<p>${fileInput.files[0].name}</p>`;
            }
        });
    }

    setupDropArea("drop-area-1", "file-input-1");
    setupDropArea("drop-area-2", "file-input-2");
});
