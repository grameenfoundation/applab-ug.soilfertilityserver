

$(document).ready(function() {

    $(".table").DataTable({
        bDestroy: true,
        dom: "lrfBtip",
        // "processing": true,"serverSide": true,
        buttons: [
            {
                extend: "pdfHtml5",

                pageSize: "A4",
                exportOptions: {
                    columns: ":visible"
                }

            }, {
                extend: "csvHtml5",
                exportOptions: {
                    columns: ":visible"
                }
            }
            //, 'colvis'
        ],
        "lengthMenu": [[10, 50, 100, 1000], [10, 50, 100, 1000]]
    });

    //$("div.toolbar").prepend('Export Visible Data As:');//for the datatables export section

    ////Format Date time picker
    //$('.date').datepicker({
    //    dateFormat: "dd/MM/yy",
    //    changeMonth: true,
    //    changeYear: true,
    //    yearRange: "-100:+0"
    //});

});