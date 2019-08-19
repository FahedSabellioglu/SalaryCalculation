$("#SalaryTypeSelect").on("change",function(){

        selectedOption = $(this).val();

        //  0 gross to net
        //  1 net to gross
        if(selectedOption==1)
        {
                $("#ConvertFrom").html("Net");
                $("#ConvertTo").html("Gross");
        }
        else if (selectedOption == 0)
        {
                $("#ConvertFrom").html("Gross");
                $("#ConvertTo").html("Net");
        }

})