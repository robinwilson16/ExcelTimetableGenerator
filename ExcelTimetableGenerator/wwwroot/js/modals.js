function doModal(title, text, size, id) {
    console.log(text);
    modalWidth = getModalWidth(size);
    modalID = getModalID(id);

    let modal = `
        <div id="${modalID}" class="modal fade" tabindex="-1" role="dialog" aria-labelledby="confirm-modal" aria-hidden="true">
            <div class="modal-dialog${modalWidth}">
                <div class="modal-content">
                    <div class="modal-header">
                    <h5 class="modal-title" id="confirm-modal"><i class="fas fa-info-circle"></i> ${title}</h5>
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                        <span aria-hidden="true">&times;</span>
                    </button>
                    </div>
                    <div class="modal-body">
                        <p>${text}</p>
                    </div>
                    <div class="modal-footer">
                        <span class="btn btn-primary" data-dismiss="modal">Close</span>
                    </div>
                </div>
            </div>
        </div>`;

    $("body").append(modal);
    $(`#${modalID}`).modal();
    $(`#${modalID}`).modal('show');

    $(`#${modalID}`).on("hidden.bs.modal", function (e) {
        $(this).remove();
    });
}

function doQuestionModal(title, text, size, id) {
    console.log(text);
    modalWidth = getModalWidth(size);
    modalID = getModalID(id);

    let modal = `
        <div id="${modalID}" class="modal fade" tabindex="-1" role="dialog" aria-labelledby="question-modal" aria-hidden="true">
            <div class="modal-dialog${modalWidth}">
                <div class="modal-content">
                    <div class="modal-header">
                    <h5 class="modal-title" id="confirm-modal"><i class="fas fa-question-circle"></i> ${title}</h5>
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                        <span aria-hidden="true">&times;</span>
                    </button>
                    </div>
                    <div class="modal-body">
                        <p>${text}</p>
                    </div>
                    <div class="modal-footer">
                        <span class="btn btn-success QuestionModal Yes" data-dismiss="modal">Yes</span>
                        <span class="btn btn-danger QuestionModal No" data-dismiss="modal">No</span>
                        <span class="btn btn-secondary QuestionModal Cancel" data-dismiss="modal">Cancel</span>
                    </div>
                </div>
            </div>
        </div>`;

    $("body").append(modal);
    $(`#${modalID}`).modal();
    $(`#${modalID}`).modal('show');

    $(`#${modalID}`).on("shown.bs.modal", function (e) {
        $(".QuestionModal").click(function (event) {
            let QuestionModalAnswerID = $("#QuestionModalAnswerID");

            if ($(this).hasClass("Yes")) {
                QuestionModalAnswerID.val("Y");
            }
            else if ($(this).hasClass("No")) {
                QuestionModalAnswerID.val("N");
            }
            else if ($(this).hasClass("Cancel")) {
                QuestionModalAnswerID.val("C");
            }
            else {
                QuestionModalAnswerID.val("X");
                doErrorModal("Invalid Option", "An invalid choice was detected. Please review your selection and try again");
            }

            doQuestionModalAction();
        });
    });

    $(`#${modalID}`).on("hidden.bs.modal", function (e) {
        $(this).remove();
    });
}

function doErrorModal(title, text, size, id) {
    console.log(text);
    modalWidth = getModalWidth(size);
    modalID = getModalID(id);

    let modal = `
        <div id="${modalID}" class="modal fade" tabindex="-1" role="dialog" aria-labelledby="confirm-modal" aria-hidden="true">
            <div class="modal-dialog${modalWidth}">
                <div class="modal-content">
                    <div class="modal-header">
                    <h5 class="modal-title" id="confirm-modal"><i class="fas fa-info-circle"></i> ${title}</h5>
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                        <span aria-hidden="true">&times;</span>
                    </button>
                    </div>
                    <div class="modal-body">
                        <p>An unexpected error has occurred which could indicate a defect with the system</p>
                        <div class="alert alert-danger" role="alert">
                        <i class="fas fa-bug"></i> ${text}
                    </div>
                    <div class="modal-footer">
                        <span class="btn btn-primary" data-dismiss="modal">Close</span>
                    </div>
                </div>
            </div>
        </div>`;

    $("body").append(modal);
    $(`#${modalID}`).modal();
    $(`#${modalID}`).modal('show');

    $(`#${modalID}`).on("hidden.bs.modal", function (e) {
        $(this).remove();
    });

    var audio = new Audio("/sounds/error.wav");
    audio.play();
}

function doCrashModal(error, size, id) {
    var stackError = $(error.responseText).find(".stackerror").html() || "Unknown error";
    var stackTrace = $(error.responseText).find(".rawExceptionStackTrace").html() || "";
    console.log(stackTrace);
    modalWidth = getModalWidth(size);
    modalID = getModalID(id);

    let modal = `
        <div id="${modalID}" class="modal fade" tabindex="-1" role="dialog" aria-labelledby="confirm-modal" aria-hidden="true">
            <div class="modal-dialog${modalWidth}">
                <div class="modal-content">
                    <div class="modal-header">
                    <h5 class="modal-title" id="confirm-modal"><i class="fas fa-info-circle"></i> ${title}</h5>
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                        <span aria-hidden="true">&times;</span>
                    </button>
                    </div>
                    <div class="modal-body">
                        <p>An unexpected error has occurred which could indicate a defect with the system</p>
                        <div class="alert alert-danger" role="alert">
                            <i class="fas fa-bug"></i> ${stackError}
                        </div>
                        <div class="pre-scrollable small">
                            <p><code>${stackTrace}</code></p>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <span class="btn btn-primary" data-dismiss="modal">Close</span>
                    </div>
                </div>
            </div>
        </div>`;

    $("body").append(html);
    $(`#${modalID}`).modal();
    $(`#${modalID}`).modal('show');

    $(`#${modalID}`).on("hidden.bs.modal", function (e) {
        $(this).remove();
    });

    var audio = new Audio("/sounds/error.wav");
    audio.play();
}

function getModalWidth(size) {
    if (size == null) {
        modalWidth = "";
    }
    else if (size === "sm") {
        modalWidth = " modal-sm";
    }
    else if (size === "lg") {
        modalWidth = " modal-lg";
    }
    else if (size === "xl") {
        modalWidth = " modal-xl";
    }
    else {
        modalWidth = "";
    }

    return modalWidth;
}

function getModalID(id) {
    if (id == null) {
        modalID = `dynamicModal`;
    }
    else {
        modalID = id;
    }

    return modalID;
}