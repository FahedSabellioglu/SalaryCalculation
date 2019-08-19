$(".dropdown-item").on("click",function(){

    selectedYear = $(this).data('year');

    $.ajax({

        url:'salary/paremeter',
        data: {'year':selectedYear},
        dataType:'json',
        success: function(data){

            $("#ModalBody").html(data.html_form);
            $("#exampleModalLong").modal('show');
            },

        error: function(data){console.log(data)}

    })
   });