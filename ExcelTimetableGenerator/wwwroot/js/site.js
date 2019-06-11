﻿$("#GenerateTimetablesButton").click(function (event) {
    let academicYear = $("#AcademicYearID").val();

    if (academicYear === "") {
        let title = `Error Determining Academic Year`;
        let content = `The current academic year could not be determined. Please check the setting in the database and try again`;

        doErrorModal(title, content);
    }
    else {
        generateTimetables(academicYear, null);
    }
});

$("#GenerateOneTimetableButton").click(function (event) {
    let academicYear = $("#AcademicYearID").val();
    let course = $("#CourseCodeID").val();

    if (academicYear === "") {
        let title = `Error Determining Academic Year`;
        let content = `The current academic year could not be determined. Please check the setting in the database and try again`;

        doErrorModal(title, content);
    }
    else if (course === "") {
        let title = `Course Not Selected`;
        let content = `Please select or type in a course first in the box above.`;

        doModal(title, content);
    }
    else {
        generateTimetables(academicYear, course);
    }
});

$(".DownloadButton").click(function (event) {
    let zipFileLocation = "/Output/Timetables.zip";
    $("#DownloadZipFrame").attr("src", zipFileLocation);
});

function generateTimetables(academicYear, courseCode) {
    return new Promise(resolve => {
        let dataToLoad = `/ExportTimetables/?academicYear=${academicYear}`;

        if (courseCode !== null) {
            dataToLoad += `&course=${courseCode}`;
        }

        var generatingMsg = `
            <div class="container">
                <div class="row">
                    <div class="col-md-3">
                        <h1 class="LoadingAnim"><i class="fas fa-spinner fa-spin"></i></h1>
                    </div>
                    <div class="col-md-9">
                        <p>
                            Timetables are now being generated.
                        </p>
                        <p>
                            You will be notified when the process has completed
                        </p>
                    </div>
                </div>
            </div>
            <div class="alert alert-primary" role="alert">
                <i class="fas fa-info-circle"></i> Please do not refresh the page and note that generation may continue even when the page is closed.
            </div>`;

        doModal(
            "Generating Timetables",
            generatingMsg,
            null,
            "GenerationInProgressDialog"
        );

        $.get(dataToLoad, function (data) {

        })
            .then(data => {
                let ttb = JSON.parse(data);
                $("#NumFilesGenerated").html(ttb.timetables.numFilesExported);
                $("#GenerationInProgressDialog").modal("hide");
                $("#GenerationCompleteModal").modal();
                //doModal(
                //    "Generation Complete",
                //    `Timetables have been successfully generated: 
                //    <ul>
                //        <li>Files Generated: ${ttb.timetables.numFilesExported}</li>
                //        <li>Save Path: ${ttb.timetables.savePath}</li>
                //    </ul>`
                //);
            })
            .fail(function () {
                let title = `Error Generating Timetables`;
                let content = `Sorry an error occurred generating the files. Please try again.`;

                doErrorModal(title, content);
            });
    });
}

$("#AboutSystemLink").click(function (event) {
    let dataToLoad = `https://raw.githubusercontent.com/robinwilson16/ExcelTimetableGenerator/master/README.md`;
    let title = `About Excel Timetable Generator`;

    $.get(dataToLoad, function (data) {

    })
        .then(data => {
            var markdown = marked(data);
            let content = `
                <p>WLC Progressions System &copy; Ealing and Hammersmith West London College</p>
                <div class="scrollable">${markdown}</div>`;

            doModal(title, content, "lg", "AboutInfo");
        })
        .fail(function () {
            let content = `Error loading content`;

            doErrorModal(title, content);
        });
    doModal(title, content);
});

$("#ChangelogLink").click(function (event) {
    let dataToLoad = `https://raw.githubusercontent.com/robinwilson16/ExcelTimetableGenerator/master/CHANGELOG.md`;
    let title = `Changelog for Excel Timetable Generator`;

    $.get(dataToLoad, function (data) {

    })
        .then(data => {
            var markdown = marked(data);

            doModal(title, markdown, "lg", "ChangelogInfo");
        })
        .fail(function () {
            let content = `Error loading content`;

            doErrorModal(title, content);
        });
});

//Old
$(function () {
    $('#btnUpload').on('click', function () {
        var fileExtension = ['xls', 'xlsx'];
        var filename = $('#fUpload').val();
        if (filename.length === 0) {
            alert("Please select a file.");
            return false;
        }
        else {
            var extension = filename.replace(/^.*\./, '');
            if ($.inArray(extension, fileExtension) === -1) {
                alert("Please select only excel files.");
                return false;
            }
        }
        var fdata = new FormData();
        var fileUpload = $("#fUpload").get(0);
        var files = fileUpload.files;
        fdata.append(files[0].name, files[0]);
        $.ajax({
            type: "POST",
            url: "/ImportExport?handler=Import",
            beforeSend: function (xhr) {
                xhr.setRequestHeader("RequestVerificationToken",
                    $('input:hidden[name="__RequestVerificationToken"]').val());
            },
            data: fdata,
            contentType: false,
            processData: false,
            success: function (response) {
                if (response.length === 0)
                    alert('Some error occured while uploading');
                else {
                    $('#dvData').html(response);
                }
            },
            error: function (e) {
                $('#dvData').html(e.responseText);
            }
        });
    });
});