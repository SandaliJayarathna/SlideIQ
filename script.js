function startCreating() {
    document.getElementById("homeScreen").classList.add("hidden");
    document.getElementById("typeScreen").classList.remove("hidden");
}
function selectType(type) {

    if (type === "Office Farewell") {

        document.getElementById("typeScreen").classList.add("hidden");

        document.getElementById("farewellScreen").classList.remove("hidden");

    } else {

        alert("This presentation type will be available soon.");

    }
}
function selectTheme(theme) {

    document.getElementById("selectedTheme").value = theme;

    const buttons = document.querySelectorAll(".theme-options button");

    buttons.forEach(function(button) {
        button.classList.remove("selected");
    });

    const selectedButton = document.querySelector(
        '.theme-options button[onclick="selectTheme(\'' + theme + '\')"]'
    );

    if (selectedButton) {
        selectedButton.classList.add("selected");
    }
}

async function generatePresentation() {

    const name = document.getElementById("personName").value;
    const position = document.getElementById("jobPosition").value;
    const company = document.getElementById("companyName").value;
    const years = document.getElementById("yearsWorked").value;
    const achievements = document.getElementById("achievements").value;
    const memories = document.getElementById("memories").value;
    const message = document.getElementById("farewellMessage").value;

    if (name === "") {
        alert("Please enter the person's name.");
        return;
    }

    const photoInput = document.getElementById("photos");
    const files = Array.from(photoInput.files);

    if (files.length > 6) {
        alert("Please select a maximum of 6 photos.");
        return;
    }

    // Show text in preview

    document.getElementById("previewName").textContent = name;
    document.getElementById("previewPosition").textContent = position;
    document.getElementById("previewCompany").textContent = company;
    document.getElementById("previewYears").textContent = years;
    document.getElementById("previewAchievements").textContent = achievements;
    document.getElementById("previewMemories").textContent = memories;
    document.getElementById("previewMessage").textContent = message;

    // Show photos in preview

    const photoPreview = document.getElementById("photoPreview");

    photoPreview.innerHTML = "";

    for (let i = 0; i < files.length; i++) {

        const image = document.createElement("img");

        image.src = URL.createObjectURL(files[i]);

        photoPreview.appendChild(image);

    }

    // Hide form and show preview

    document.getElementById("farewellScreen").classList.add("hidden");

    document.getElementById("previewScreen").classList.remove("hidden");

    // Apply theme

    const theme = document.getElementById("selectedTheme").value;

    const slides = document.querySelectorAll(".slide");

    slides.forEach(function(slide) {

        slide.classList.remove(
            "theme-elegant",
            "theme-professional",
            "theme-celebration",
            "theme-simple"
        );

        slide.classList.add("theme-" + theme);

    });
        // Convert photos to Base64

    let photos = [];

    try {

        for (let i = 0; i < files.length; i++) {

            const reader = new FileReader();

            const base64Image = await new Promise(function(resolve, reject) {

                reader.onload = function() {
                    resolve(reader.result);
                };

                reader.onerror = function() {
                    reject(reader.error);
                };

                reader.readAsDataURL(files[i]);

            });

            photos.push(base64Image);

        }

    } catch (error) {

        console.error(error);

        alert("Could not read the photos.");

        return;

    }


    // Send data to server

    try {

        const response = await fetch("/generate-pptx", {

            method: "POST",

            headers: {
                "Content-Type": "application/json"
            },

            body: JSON.stringify({

                personName: name,
                jobPosition: position,
                companyName: company,
                yearsWorked: years,
                achievements: achievements,
                memories: memories,
                farewellMessage: message,
                photos: photos

            })

        });


        if (!response.ok) {

            throw new Error("PowerPoint generation failed.");

        }


        // Receive PowerPoint file

        const blob = await response.blob();


        // Download PowerPoint

        //const downloadUrl = window.URL.createObjectURL(blob);

        //const link = document.createElement("a");

        //link.href = downloadUrl;

       //link.download = "office-farewell-presentation.pptx";

       //document.body.appendChild(link);

        //link.click();

        //link.remove();

        //window.URL.revokeObjectURL(downloadUrl);


        alert("PowerPoint presentation created successfully!");

    } catch (error) {

        console.error(error);

        alert("Could not create PowerPoint presentation.");

    }

}
const photoInput = document.getElementById("photos");

if (photoInput) {
    photoInput.addEventListener("change", function () {

        const count = this.files.length;

        const text = count === 0
            ? "No photos selected"
            : count + (count === 1 ? " photo selected" : " photos selected");

        document.getElementById("selectedPhotoCount").textContent = text;

    });
}
