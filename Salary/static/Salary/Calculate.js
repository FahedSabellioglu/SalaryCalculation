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
            var can_pass = true
            $('td input').each(function(){

                if(parseFloat($(this).val())<2558.4)
                {
                    can_pass = false
                    alert("The min value you can enter is 2558.4")
                    return false
                }

                else if($(this).val()!="")
                {
                    values[$(this).attr('name')] = parseFloat($(this).val());
                }
            });

            if (can_pass==false)
            {
                return false
            }

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