normalize_text = function (text) {

    //remove special characters
    text = text.replace(/([^\u0621-\u063A\u0641-\u064A\u0660-\u0669a-zA-Z 0-9])/g, '');

    //normalize Arabic
    text = text.replace(/(آ|إ|أ)/g, 'ا');
    text = text.replace(/(ة)/g, 'ه');
    // text = text.replace(/(ئ|ؤ)/g, 'ء')
    text = text.replace(/(ى)/g, 'ي');
    text = text.replace(/(عبد ال)/g, 'عبدال');

    //convert arabic numerals to english counterparts.
    var starter = 0x660;
    for (var i = 0; i < 10; i++) {
        text.replace(String.fromCharCode(starter + i), String.fromCharCode(48 + i));
    }

    return text;
}

// Apply normalize_text() to all properties of the object
preprocess = function(object){
    for (var key in object) {  
        object[key] = normalize_text(object[key])
    }
    return object;
}

function init_autoCompleteJS(json_object) {

    // The autoComplete.js Engine instance creator
    const autoCompleteJS = new autoComplete({
        data: {
            src: async () => {
                try {
                    // Loading placeholder text
                    document
                        .getElementById("autoComplete")
                        .setAttribute("placeholder", "Loading...");
                    // Fetch External Data Source
                    // const source = await fetch(
                    //     "https://tarekraafat.github.io/autoComplete.js/demo/db/generic.json"
                    // );
                    // const data = await source.json();
                    var json = JSON.parse(json_object);
                    var preprocessed = json.map(element => preprocess(element));
                    const data = await preprocessed;
                    // Post Loading placeholder text
                    document
                        .getElementById("autoComplete")
                        .setAttribute("placeholder", autoCompleteJS.placeHolder);
                    // Returns Fetched data
                    return data;
                } catch (error) {
                    return error;
                }
            },
            keys: ["الاسم", "الكود"],
            cache: true,
            filter: (list) => {
                // Filter duplicates
                // incase of multiple data keys usage
                const filteredResults = Array.from(
                    new Set(list.map((value) => value.match))
                ).map((food) => {
                    return list.find((value) => value.match === food);
                });

                return filteredResults;
            }
        },
        placeHolder: "Search for Student ID & Names!",
        resultsList: {
            element: (list, data) => {
                const info = document.createElement("p");
                if (data.results.length > 0) {
                    info.innerHTML = `Displaying <strong>${data.results.length}</strong> out of <strong>${data.matches.length}</strong> results`;
                } else {
                    info.innerHTML = `Found <strong>${data.matches.length}</strong> matching results for <strong>"${data.query}"</strong>`;
                }
                list.prepend(info);
            },
            noResults: true,
            maxResults: 15,
            tabSelect: true
        },
        resultItem: {
            element: (item, data) => {
                // Modify Results Item Style
                item.style = "display: flex; justify-content: space-between;";
                // Modify Results Item Content
                item.innerHTML = `
      <span style="text-overflow: ellipsis; white-space: nowrap; overflow: hidden;">
        ${data.value['الاسم']} ${data.value['الكود']}
      </span>
      <span style="display: flex; align-items: center; font-size: 13px; font-weight: 100; text-transform: uppercase; color: rgba(0,0,0,.2);">
        ${data.key}
      </span>`;
            },
            highlight: true
        },
        events: {
            input: {
                focus: () => {
                    if (autoCompleteJS.input.value.length) autoCompleteJS.start();
                }
            }
        }
    });

    // autoCompleteJS.input.addEventListener("init", function (event) {
    //   console.log(event);
    // });

    // autoCompleteJS.input.addEventListener("response", function (event) {
    //   console.log(event.detail);
    // });

    // autoCompleteJS.input.addEventListener("results", function (event) {
    //   console.log(event.detail);
    // });

    // autoCompleteJS.input.addEventListener("open", function (event) {
    //   console.log(event.detail);
    // });

    // autoCompleteJS.input.addEventListener("navigate", function (event) {
    //   console.log(event.detail);
    // });

    autoCompleteJS.input.addEventListener("selection", function (event) {
        $('#input').val('');
        $(".selection").html('');
        $("#savebtn").attr("disabled", true);
        $("#exportbtn").attr("disabled", true);

        const feedback = event.detail;
        autoCompleteJS.input.blur();
        // Prepare User's Selected Value
        const selection = feedback.selection.value[feedback.selection.key];
        const StudentName = feedback.selection.value['الاسم'];
        const StudentID = feedback.selection.value['الكود'];

        window.StudentID = StudentID;
        // Render selected choice to selection div
        document.querySelector(".selection").innerHTML = StudentName + '  ' + StudentID;
        // Replace Input value with the selected value
        autoCompleteJS.input.value = selection;
        console.log(StudentName);
        // console.log(feedback);

    });

    // autoCompleteJS.input.addEventListener("close", function (event) {
    //   console.log(event.detail);
    // });

    // Toggle Search Engine Type/Mode
    document.querySelector(".toggler").addEventListener("click", () => {
        // Holds the toggle button selection/alignment
        const toggle = document.querySelector(".toggle").style.justifyContent;

        if (toggle === "flex-start" || toggle === "") {
            // Set Search Engine mode to Loose
            document.querySelector(".toggle").style.justifyContent = "flex-end";
            document.querySelector(".toggler").innerHTML = "Loose";
            autoCompleteJS.searchEngine = "loose";
        } else {
            // Set Search Engine mode to Strict
            document.querySelector(".toggle").style.justifyContent = "flex-start";
            document.querySelector(".toggler").innerHTML = "Strict";
            autoCompleteJS.searchEngine = "strict";
        }
    });

    // Blur/unBlur page elements
    const action = (action) => {
        const title = document.querySelector("h1");
        const mode = document.querySelector(".mode");
        const selection = document.querySelector(".selection");
        const footer = document.querySelector(".footer");

        if (action === "dim") {
            title.style.opacity = 1;
            mode.style.opacity = 1;
            selection.style.opacity = 1;
        } else {
            title.style.opacity = 0.3;
            mode.style.opacity = 0.2;
            selection.style.opacity = 0.1;
        }
    };

    // Blur/unBlur page elements on input focus
    ["focus", "blur"].forEach((eventType) => {
        autoCompleteJS.input.addEventListener(eventType, () => {
            // Blur page elements
            if (eventType === "blur") {
                action("dim");
            } else if (eventType === "focus") {
                // unBlur page elements
                action("light");
            }
        });
    });
}


var ExcelToJSON = function () {

    this.parseExcel = function (file) {
        var reader = new FileReader();

        reader.onload = function (e) {
            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });
            workbook.SheetNames.forEach(function (sheetName) {
                // Here is your object
                var excel_data = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName], { range: 5 });
                for(var i = 0; i < 5; i++){
                    excel_data.unshift(excel_data[0]);
                }                
                window.excel_data = excel_data;
                var json_object = JSON.stringify(window.excel_data);
                // jQuery( '#xlx_json' ).val( json_object );
                init_autoCompleteJS(json_object);

                // console.log(window.excel_data);

            })
        };

        reader.onerror = function (ex) {
            console.log(ex);
        };

        reader.readAsBinaryString(file);
    };
};

function handleFileSelect(evt) {

    var files = evt.target.files; // FileList object
    var xl2json = new ExcelToJSON();
    xl2json.parseExcel(files[0]);
    $('#upload').hide();
    window.filename = files[0].name
    $('#title').text("Type a new column name to add marks in : ");
    $('#col_name_input').show();
}

document.getElementById('upload').addEventListener('change', handleFileSelect, false);

$('#input').on('keyup', function () {
    $("#savebtn").attr("disabled", false);
})

function Save() {

    for (var i = 0; i < window.excel_data.length; i++) {
        if (window.excel_data[i]['الكود'] == window.StudentID) {
            window.excel_data[i][window.col_name] = $('#input').val();
            // log student mark
            console.log(window.col_name + " : "+$('#input').val())
            break;
        }
    }

    $("#exportbtn").attr("disabled", false);
    $(".selection").html($(".selection").html() + '<p>' + window.col_name + ' ' + $('#input').val() + '</p>');
    $('#input').val('');
    $("#savebtn").attr("disabled", true);
}

function showControls() {
    $('#col_name_input').hide();
    $('#title').text("Search in : "+window.filename);
    $('#controls').show();
    window.col_name = $('#col_name').val()
    $('#label_col').text(window.col_name + " :")
}


function ExportData() {
    // var wb = XLSX.utils.book_new();

    // To Enable RTL workbook change RTL: false to  RTL: true
    var wb = { Workbook: { Views: [{ RTL: false }] }, Sheets: {}, SheetNames: [] }

    var ws = XLSX.utils.json_to_sheet(window.excel_data);

    var today = new Date();
    var date = today.getFullYear() + '-' + (today.getMonth() + 1) + '-' + today.getDate();
    var time = today.getHours() + "-" + today.getMinutes() + "-" + today.getSeconds();
    var dateTime = date + '__' + time;

    XLSX.utils.book_append_sheet(wb, ws, "snap @ " + dateTime);
    XLSX.writeFile(wb, "output_" + dateTime + ".xlsx");
}
