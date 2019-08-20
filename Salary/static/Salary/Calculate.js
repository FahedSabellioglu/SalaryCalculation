$("#CalBtn").on("click",function(){


//        <!-- Users Parameters -->
            calType = parseInt($("#SalaryTypeSelect").val());
            socialStatus = parseInt($("#SocialStatusSelect").val());
            year = parseInt($("#YearSelect").val());
            empCost = $("#empCost").prop("checked") ;
            empSocialShare = $("#empSocialShare").prop("checked");
            partnerStatus = parseInt($("#PartnerStatus").val());
            kidsCount = parseInt($("#ChildrensCount").val());
//        <!--         -->

        //        <!-- Salaries Control -->

            var values = {};
            $('td input').each(function(){
                if($(this).val()!="")
                {
                    values[$(this).attr('name')] = parseFloat($(this).val());
                }
            });

            if(Object.keys(values).length==0)
            {
                alert("Please fill at least one field.")
                return false
            }
            else if (Object.keys(values).length!=0)
            {
                usrParameters={'calType':calType,'socialStatus':socialStatus,
                           'year':year,'empSocialShare':empSocialShare ,'empCost':empCost,
                           'partnerStatus':partnerStatus,'kidsCount':kidsCount,'salaries':JSON.stringify(values)
                       }

                if(socialStatus == 1)
                {
                    usrParameters['partnerStatus'] = partnerStatus;
                    usrParameters['kidsCount'] = kidsCount;
                }

                $.ajax({

                    url:'calculation',
                    data:usrParameters,
                    dataType:'json',
                    success: function(data){
                        $("#tableBody").html(data.html_form);
                        $("#imgDiv").show();
                    },
                    error: function(){
                        alert("Something went wrong, Please try again later.")
                        }

                });

            }

        //        <!--         -->

});