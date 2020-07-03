$("#CourseCodeID").keyup(function (event) {
    if (event.keyCode == 13) {
        $("#GenerateOneTimetableButton").trigger("click");
    }
});

$("#GenerateTimetablesButton").click(function (event) {
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
        let dataToLoad = `/Programmes/Details/${course}/?handler=Json`;

        let title = `Checking Course Data`;
        let content = `
            <div class="container">
                <div class="row">
                    <div class="col-md-3">
                        <h1 class="LoadingAnim"><i class="fas fa-spinner fa-spin"></i></h1>
                    </div>
                    <div class="col-md-9">
                        <p>
                            Please wait whilst course data is validated
                        </p>
                    </div>
                </div>
            </div>`;

        doModal(
            title,
            content,
            null,
            "CheckingCourseDialog"
        );

        $.get(dataToLoad, function (data) {

        })
            .then(data => {
                console.log(dataToLoad + " Loaded");

                $("#CheckingCourseDialog").modal("hide");

                if (data.programme != null) {
                    generateTimetables(academicYear, course);
                }
                else {
                    let title = `Course Invalid`;
                    let content = `The course code you entered is not a valid course code: "${course}".<br />Please try again.`;

                    doModal(title, content);
                    $("#CourseCodeID").val("");
                }
            });
    }
});

$(".DownloadButton").click(function (event) {
    let sessionID = $("#SessionID").val();
    url = `/Exports/${sessionID}/Timetables.zip`;
    event.originalEvent.currentTarget.href = url;

    $('#GenerationCompleteModal').modal('hide')
});

function generateTimetables(academicYear, courseCode) {
    return new Promise(resolve => {
        let sessionID = $("#SessionID").val();

        if (sessionID == null) {
            let title = `Error Generating Timetables`;
            let content = `Sorry the system was unable to determine your session ID.`;

            doErrorModal(title, content);

            return;
        }

        let dataToLoad = `/ExportTimetables/${sessionID}/?academicYear=${academicYear}`;

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
                let haveReadPermission = (ttb.timetables.haveReadPermission == "True");
                let haveWritePermission = (ttb.timetables.haveWritePermission == "True");
                let outputPath = ttb.timetables.outputPath;
                $("#NumFilesGenerated").html(ttb.timetables.numFilesExported);
                $("#GenerationInProgressDialog").modal("hide");

                if (haveReadPermission === false) {
                    let title = `Error Generating Timetables`;
                    let content = `
                        Sorry an error occurred generating the files as this web application does not have permission to read data in the following folder:
                        <div class="pre-scrollable small">
                            <p><code>${outputPath}</code></p>
                        </div>
                        `;

                    doErrorModal(title, content, "lg");
                }
                else if (haveWritePermission === false) {
                    let title = `Error Generating Timetables`;
                    let content = `
                        Sorry an error occurred generating the files as this web application does not have permission to write to the following folder:
                        <div class="pre-scrollable small">
                            <p><code>${outputPath}</code></p>
                        </div>
                        `;

                    doErrorModal(title, content, "lg");
                }
                else {
                    $("#GenerationCompleteModal").modal();
                    $("#CourseCodeID").val("");
                }
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