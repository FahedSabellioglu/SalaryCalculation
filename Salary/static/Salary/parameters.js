$(".dropdown-item").on("click",function(){

    selectedYear = $(this).data('year');

    $.ajax({

        url:'parameter',
        data: {'year':selectedYear},
        dataType:'json',
        success: function(data){

            $("#ModalBody").html(data.html_form);
            $("#exampleModalLong").modal('show');
            },

        error: function(data){

                alert("Something went wrong, Please try again later")
            }

    })
   });