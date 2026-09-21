// ===============================
// SCREEN CONTROL
// ===============================

function showScreen(screenId) {

    const screens = [
        "homeScreen",
        "typeScreen",
        "farewellScreen",
        "previewScreen"
    ];

    screens.forEach(function(id) {

        const screen = document.getElementById(id);

        if (screen) {
            screen.classList.add("hidden");
        }

    });

    const selectedScreen = document.getElementById(screenId);

    if (selectedScreen) {
        selectedScreen.classList.remove("hidden");
    }

    window.scrollTo({
        top: 0,
        behavior: "smooth"
    });
}


// ===============================
// START
// ===============================

function startCreating() {

    showScreen("typeScreen");

}


// ===============================
// SELECT TYPE
// ===============================

function selectType(type) {

    if (type === "Office Farewell") {

        showScreen("farewellScreen");

    } else {

        alert("This presentation type will be available soon.");

    }

}


// ===============================
// SELECT THEME
// ===============================

function selectTheme(theme) {

    const selectedTheme =
        document.getElementById("selectedTheme");

    if (selectedTheme) {
        selectedTheme.value = theme;
    }

    const buttons =
        document.querySelectorAll(".theme-button");

    buttons.forEach(function(button) {

        button.classList.remove("selected");

    });

    const selectedButton =
        document.querySelector(
            '.theme-button[onclick="selectTheme(\'' + theme + '\')"]'
        );

    if (selectedButton) {
        selectedButton.classList.add("selected");
    }

    applyTheme(theme);
}


// ===============================
// APPLY THEME
// ===============================

function applyTheme(theme) {

    const liveSlide =
        document.getElementById("liveSlide");

    if (liveSlide) {

        liveSlide.classList.remove(
            "theme-elegant",
            "theme-professional",
            "theme-celebration",
            "theme-simple"
        );

        liveSlide.classList.add(
            "theme-" + theme
        );

    }

    const coverSlides =
        document.querySelectorAll(".cover-slide");

    coverSlides.forEach(function(slide) {

        slide.classList.remove(
            "theme-elegant",
            "theme-professional",
            "theme-celebration",
            "theme-simple"
        );

        slide.classList.add(
            "theme-" + theme
        );

    });
}


// ===============================
// LIVE PREVIEW
// ===============================

function updateLivePreview() {

    const name =
        document.getElementById("personName").value;

    const position =
        document.getElementById("jobPosition").value;

    const company =
        document.getElementById("companyName").value;


    document.getElementById("previewName").textContent =
        name || "Person Name";

    document.getElementById("previewPosition").textContent =
        position || "Job Position";

    document.getElementById("previewCompany").textContent =
        company || "Company / Organization";
}


// ===============================
// GENERATE PRESENTATION
// ===============================

async function generatePresentation() {

    const name =
        document.getElementById("personName").value.trim();

    const position =
        document.getElementById("jobPosition").value.trim();

    const company =
        document.getElementById("companyName").value.trim();

    const years =
        document.getElementById("yearsWorked").value.trim();

    const achievements =
        document.getElementById("achievements").value.trim();

    const memories =
        document.getElementById("memories").value.trim();

    const message =
        document.getElementById("farewellMessage").value.trim();


    if (name === "") {

        alert("Please enter the person's name.");

        return;

    }


    const photoInput =
        document.getElementById("photos");

    const files =
        Array.from(photoInput.files);


    if (files.length > 6) {

        alert("Please select a maximum of 6 photos.");

        return;

    }


    // Update final preview

    document.getElementById("finalPreviewName").textContent =
        name;

    document.getElementById("finalPreviewPosition").textContent =
        position || "Job Position";

    document.getElementById("finalPreviewCompany").textContent =
        company || "Company / Organization";

    document.getElementById("previewYears").textContent =
        years || "0";

    document.getElementById("previewAchievements").textContent =
        achievements || "No achievements added.";

    document.getElementById("previewMemories").textContent =
        memories || "No special memories added.";

    document.getElementById("previewMessage").textContent =
        message || "Wishing you all the best for the future!";


    // Photos

    const photoPreview =
        document.getElementById("photoPreview");

    photoPreview.innerHTML = "";


    for (let i = 0; i < files.length; i++) {

        const image =
            document.createElement("img");

        image.src =
            URL.createObjectURL(files[i]);

        photoPreview.appendChild(image);

    }


    if (files.length === 0) {

        photoPreview.innerHTML =
            '<div style="grid-column:1/-1;color:#999;font-size:11px;">No photos added.</div>';

    }


    // Theme

    const theme =
        document.getElementById("selectedTheme").value;

    applyTheme(theme);


    // Show preview

    showScreen("previewScreen");


    // Convert photos

    let photos = [];


    try {

        for (let i = 0; i < files.length; i++) {

            const reader =
                new FileReader();

            const base64Image =
                await new Promise(function(resolve, reject) {

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


    // Send to server

    try {

        const response =
            await fetch("/generate-pptx", {

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

                    photos: photos,

                    theme: theme

                })

            });


        if (!response.ok) {

            throw new Error(
                "PowerPoint generation failed."
            );

        }


        const blob =
            await response.blob();


        downloadBlob(blob);


        alert(
            "PowerPoint presentation created successfully!"
        );


    } catch (error) {

        console.error(error);

        alert(
            "Could not create PowerPoint presentation."
        );

    }

}


// ===============================
// DOWNLOAD
// ===============================

let lastGeneratedBlob = null;


function downloadBlob(blob) {

    lastGeneratedBlob = blob;

    const downloadUrl =
        window.URL.createObjectURL(blob);

    const link =
        document.createElement("a");

    link.href = downloadUrl;

    link.download =
        "office-farewell-presentation.pptx";

    document.body.appendChild(link);

    link.click();

    link.remove();

    setTimeout(function() {

        window.URL.revokeObjectURL(downloadUrl);

    }, 1000);

}


function downloadAgain() {

    if (!lastGeneratedBlob) {

        alert(
            "Please generate the presentation first."
        );

        return;

    }

    downloadBlob(lastGeneratedBlob);

}


// ===============================
// PHOTO COUNTER
// ===============================

document.addEventListener(
    "DOMContentLoaded",
    function() {

        showScreen("homeScreen");


        const photoInput =
            document.getElementById("photos");


        if (photoInput) {

            photoInput.addEventListener(
                "change",
                function() {

                    const count =
                        this.files.length;

                    const text =
                        count === 0
                            ? "No photos selected"
                            : count +
                              (count === 1
                                  ? " photo selected"
                                  : " photos selected");


                    document.getElementById(
                        "selectedPhotoCount"
                    ).textContent = text;

                }
            );

        }


        // Live preview inputs

        const liveInputs = [
            "personName",
            "jobPosition",
            "companyName"
        ];


        liveInputs.forEach(function(id) {

            const input =
                document.getElementById(id);

            if (input) {

                input.addEventListener(
                    "input",
                    updateLivePreview
                );

            }

        });

    }
);