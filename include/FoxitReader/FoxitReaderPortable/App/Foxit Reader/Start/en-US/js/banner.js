$(function(){

	var cPDF_banner_boxWidth=$("#cPDF_banner").width();
	var cPDF_banner_boxHeight=$("#cPDF_banner").height();
	var cPDF_banner_LiWidth=$("#cPDF_banner").children("ul").children("li").eq(0).width();
	var cPDF_banner_liNubr=$("#cPDF_banner").find('li').length;
	//var cPDF_banner_Speed=8000;

	var cPDF_banner_Last_NextHeight=$("#cPDF_banner_Last").height();

	var cPDF_banner_Tab_AWidth=(1/cPDF_banner_liNubr)*100;
	//var cPDF_banner_TabWidth=$("#cPDF_banner_Tab").width();
	var cPDF_bannerWidth=$("#cPDF_banner").width();
	//$("#cPDF_banner_Tab").css("left",(cPDF_bannerWidth-cPDF_banner_TabWidth)*0.5);
	$("#cPDF_banner_Last").css("top",(cPDF_banner_boxHeight-cPDF_banner_Last_NextHeight)*0.5);
	$("#cPDF_banner_Next").css("top",(cPDF_banner_boxHeight-cPDF_banner_Last_NextHeight)*0.5);
	$("#prevL").css("left",-cPDF_banner_LiWidth);
	$("#prevR").css("right",-cPDF_banner_LiWidth);
	
	
	
	var currTab = parseInt($("#currTab").val());
	initBanner(currTab);
	initArr(currTab);
	
	

	
	//$("#cPDF_banner_Next").click(Slide_Next);
	//$("#cPDF_banner_Last").click(Slide_Last);
	$("#cPDF_banner_Next").click(function(){
		turnBannerPage("next");
	});
	$("#cPDF_banner_Last").click(function(){
		turnBannerPage("prev");
	});
	
	
	
	

	
});
	function initArr(curr_tab){
		var cPDF_banner_liNubr=$("#cPDF_banner").find('li').length;
		
		if(curr_tab == 0){						
			$("#cPDF_banner_Last").hide();
			$("#cPDF_banner_Next").show();					
		}else if(curr_tab == cPDF_banner_liNubr-1 && cPDF_banner_liNubr > 1){
			$("#cPDF_banner_Last").show();
			$("#cPDF_banner_Next").hide();			
		}else{
			$("#cPDF_banner_Next,#cPDF_banner_Last").show();			
		}
		
	}

	function initBanner(curr_tab){
		var cPDF_banner_liNubr=$("#cPDF_banner").find('li').length;
		var cPDF_banner_Tab_AWidth=(1/cPDF_banner_liNubr)*100;
		var cPDF_banner_LiWidth=$("#cPDF_banner").children("ul").children("li").eq(0).width();
		var cPDF_banner_Tab_Contne="";
		for(var i=0;i<parseInt(cPDF_banner_liNubr);i++){
			$("#cPDF_banner").children("ul").children("li").eq(i).css("left",-(cPDF_banner_liNubr-i-1-curr_tab)*cPDF_banner_LiWidth);			
			
		}		
	}
	
	

	function turnBannerPage(type){
		var currTab = parseInt($("#currTab").val());
		var cPDF_banner_liNubr=$("#cPDF_banner").find('li').length;
		var cPDF_banner_LiWidth=$("#cPDF_banner").children("ul").children("li").eq(0).width();
		if(type == "next"){
			if(currTab < cPDF_banner_liNubr-1) currTab++;
			$("#cPDF_banner_Next").hide();
			
		}else{
			if(currTab >= 1) currTab--;		
			$("#cPDF_banner_Last").hide();
			cPDF_banner_LiWidth = -cPDF_banner_LiWidth;
		}
		
		
		$("#currTab").val(currTab);
		
		for(var k=0;k<parseInt(cPDF_banner_liNubr);k++){
			var currPos = parseInt($("#cPDF_banner").children("ul").children("li").eq(k).css("left"))+ cPDF_banner_LiWidth;
			$("#cPDF_banner").children("ul").children("li").eq(k).animate({left:currPos},"slow",function(){initArr(currTab);});
		}	
		
		
	}


	
	function activeTab(tmp_li_obj, appEv){		
		if(tmp_li_obj == 'settings_tab'){
			$("li").remove(".sign_up_tab");	
			var p = 0;
			if(appEv) p = $("#currTab").val()-1;
			$("#currTab").val(p);
			
			if($("#cPDF_banner_Next").is(":hidden")){
				$('#cPDF_banner_Last').click();
			}else{
				$('#cPDF_banner_Next').click();
			}
			initBanner(0);
		}else if(tmp_li_obj == 'sign_up_tab'){
			var exist =$("#cPDF_banner").find(".sign_up_tab").length;
			if(exist == 0){
				var content = $("#banner_sign_up").html();
				$('<li class="carousel_li sign_up_tab">'+content+ '</li>').appendTo(".carousel_ul");
			}
		}
		
	}

