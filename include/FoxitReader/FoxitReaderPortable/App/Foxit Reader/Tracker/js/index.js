;(function ($,externalDispatchFun) {
    $(function () {
	    var moduleName = 'ReviewTracker';
	    var typeList = {
		    "startReview":      [21, 22, 23],
		    "joinReivew":       31,
		    "startEmailReview": 20,
		    "joinEmailReview":  30
	    };
		var data = null;//'{"End":0,"Action":[{"AcName":"ViewComments","AcTitle":"View Comments","Module":"ReviewTracker"},{"AcName":"EmailAllReviewers","AcTitle":"Email All Reviewers","Module":"ReviewTracker"},{"AcName":"ChangeDeadline","AcTitle":"Change Deadline","Module":"ReviewTracker"},{"AcName":"EndReview","AcTitle":"End Review","Module":"ReviewTracker"},{"AcName":"AddReviewers","AcTitle":"Add Reviewers","Module":"ReviewTracker"},{"AcName":"StartNewReviewWithSameReviewers","AcTitle":"Start New Review with Same Reviewers","Module":"ReviewTracker"}],"Content":[{"CTName":"FileLocation","CTTitle":"File Location:","CTValue":"C:\\Users\\Json_chen\\Desktop\\FZ02-04_review.pdf"},{"CTName":"SentTime","CTTitle":"Sent:","CTValue":"2016/12/29 16:20:09"},{"CTName":"Deadline","CTTitle":"Deadline:","CTValue":"2017/1/13 16:19:58"},{"CTName":"Status","CTTitle":"Status:","CTValue":"Active"},{"CTName":"ErrorMsg","CTTitle":"","CTValue":""}],"Title":"Reviews > Sent > FZ02-04_review","Type":21,"UserContent":{"ListStat":[{"CTName":"Comments","CTTitle":"Comments:","CTValue":"0 new / 0 total"},{"CTName":"Reviewers","CTTitle":"Reviewers:","CTValue":"0 new / 1 active"}],"ListTitle":["Email","Reviewer Name","Title","Comments: New/Total","Type"],"UCTitle":"Reviewers"}}';//'{"End":1,"Action": [{"AcName": "ViewComments","AcTitle": "View Comments","Module": "ReviewTracker"},{"AcName": "EmailAllReviewers","AcTitle": "Email All Reviewers","Module": "ReviewTracker"},{"AcName": "EmailInitiator","AcTitle": "Email Initiator","Module": "ReviewTracker"}],"Content": [{"CTName": "FileLocation","CTTitle": "File Location:","CTValue": "C:\\Users\\Json_chen\\Desktop\\LR9.5中文使用手册_A-1a_review.pdf"},{"CTName": "SentTime","CTTitle": "Recevied On:","CTValue": "2016/12/20 17:28:24"},{"CTName": "Deadline","CTTitle": "Deadline:","CTValue": "None"},{"CTName": "Status","CTTitle": "Status:","CTValue": "Active"},{"CTName": "ErrorMsg","CTTitle": "","CTValue": ""}],"Title": "Reviews > Joined > LR9.5中文使用手册_A-1a_review","Type": 31,"UserContent": {"ListStat": [{"CTName": "Comments","CTTitle": "Comments:","CTValue": "0 new / 0 total"},{"CTName": "Reviewers","CTTitle": "Reviewers:","CTValue": "0 new / 1 active"}],"ListTitle": ["Email","Reviewer Name","Title","Comments: New/Total","Type"],"UCTitle": "Reviewers"}}';
		try {
			data = externalDispatchFun(moduleName, "GetContent", '');
		} catch (ex) {}

	    if(data) {
		    try {
			    data  = JSON.parse(data);
		    }catch (e){
			    try {
				    data = eval("(" + data + ")");
			    }catch (ex){
				    return;
			    }
		    }
		    try
		    {
			    var type = data ? parseInt(data.Type) : 0;

			    parseData(type);

			    $('.content').show();
		    }catch (ex){
			    alert(ex.message);
			    return;
		    }
	    }

	    function parseData(type){
		    var content = data.Content ? data.Content:[];
		    var title = data.Title ? data.Title : '';
		    var action = data.Action ? data.Action : [];
		    var userContent = data.UserContent ? data.UserContent : {};
		    var utcTitle = data.UCTitle ? data.UCTitle :'';
		    var end = data.End ? parseInt(data.End) : '';
            var $loading = $('.loading');
		    $loading.attr('src',folderName+'loading.gif');

		    $('#title').text(title);
		    $('#viewComments').text(utcTitle);
		    var pageContent = '';
		    $.each(content,function(index,value){
			    var ddStr = value.CTName == 'FileLocation'?'<span class="break-word a" name="viewComments">'+value.CTValue+'</span>': (value.CTName == 'Deadline' || (value.CTName == 'Status' && end != 1)) && $.inArray(type,typeList.startReview) != -1 ?'<div>'+value.CTValue+'</div>': value.CTValue;
				var ddClass = (value.CTName == 'Deadline' || (value.CTName == 'Status' && end != 1)) && $.inArray(type,typeList.startReview) != -1 ? 'class="clearfix po-re"' : '';
			    var addStyle = value.CTName == 'ErrorMsg' ? 'style="color:red;line-height: 1.4"' : '';
			    pageContent += '<dt>'+value.CTTitle+'</dt><dd id="'+value.CTName+'" '+ddClass+' '+addStyle+'>'+ddStr+'</dd>'
		    });
		    $('#content').html(pageContent);
			$('#reviewers').html('<div class="caption-wrap" >'+userContent.UCTitle+'</div>');

		    var pageFooter = '';
		    $.each(userContent.ListStat,function(index,value){
			    pageFooter += '<dt>'+value.CTTitle+'</dt><dd id="'+value.CTName+'">'+value.CTValue+'</dd>';
		    });
            $('#total').html(pageFooter);

		    var listTr = '';
		    $.each(userContent.ListTitle,function(index,value){
			    listTr += '<th>'+value+'</th>';
		    });
		    $('#listTr').html(listTr);


		    $('#footer').html('');
		    $.each(action,function(index,value){
			    var acName = value.AcName;
			    var src = '';
			    var actionStr = '<div name="'+acName+'"> <img src="'+src+'" alt=""/> <span class="a" id="'+acName+'">'+value.AcTitle+'</span></div>';

			    switch (acName){
				    case 'EmailInitiator':
					    src= folderName + "email_initiator.png";
					    $('#footer').append(actionStr).find('div[name="'+acName+'"]').find('img').attr('src',src);
						$('#footer').find('span#'+acName).on('click',emailInitiator);
					    break;
				    case 'EmailAllReviewers':
					    src= folderName + "email_all.png";
					    $('#footer').append(actionStr).find('div[name="'+acName+'"]').find('img').attr('src',src);
					    $('#footer').find('span#'+acName).on('click',emailAllReviewers);
					    break;
				    case 'StartNewReviewWithSameReviewers':
					    src= folderName + "start_same_new_review.png";
					    $('#footer').append(actionStr).find('div[name="'+acName+'"]').find('img').attr('src',src);
					    $('#footer').find('span#'+acName).on('click',startNewReviewWithSameReviewers);
					    break;
				    case 'ChangeDeadline':
					    src = folderName + "change_deadline.png";
					    $('#Deadline').append(actionStr).find('img').attr('src',src);
					    $('#Deadline').find('span').on('click',changeDeadline);
					    break;
				    case 'EndReview':
					    src = folderName + "end_review.png";
					    $('#Status').append(actionStr).find('img').attr('src',src);
					    $('#Status').find('span').on('click',endReview);
					    break;
				    case 'AddReviewers':
					    src = folderName + "add_reviewers.png";
				        var str = '<div class="caption-info">'+actionStr+'</div>';
					    $('#reviewers').append(str).find('img').attr('src',src);
					    $('#reviewers').find('span').on('click',addReviewers);
					    break;
				    case 'ViewComments':
					    $('[name="viewComments"]').on('click',viewComments);
					    break;
			    }
		    });
	    }


	    function changeDeadline(){
		    try {
			   externalDispatchFun(moduleName, "ChangeDeadline", '');
		    } catch (ex) {}
	    }

	    function endReview(){
		    try {
			    externalDispatchFun(moduleName, "EndReview", '');
		    } catch (ex) {}
	    }

	    function addReviewers(){
		    try {
			    externalDispatchFun(moduleName, "AddReviewers", '');
		    } catch (ex) {}
	    }

	    function emailAllReviewers(){
		    try {
			    externalDispatchFun(moduleName, "EmailAllReviewers", '');
		    } catch (ex) {}
	    }

	    function startNewReviewWithSameReviewers(){
		    try {
			    externalDispatchFun(moduleName, "StartNewReviewWithSameReviewers", '');
		    } catch (ex) {}
	    }

	    function emailInitiator(){
		    try {
			    externalDispatchFun(moduleName, "EmailInitiator", '');
		    } catch (ex) {}
	    }

	    function viewComments(){
		    try {
			    externalDispatchFun(moduleName, "ViewComments", '');
		    } catch (ex) {}
	    }

    });

})(jQuery,externalDispatchFun);

var folderName = '';
function parseList(userList){
	try {
		userList = JSON.parse(userList);
	}catch (e){
		userList = eval("(" + userList + ")");
	}
	if(!userList) return;
	try {
		var trStr = '';
		var userContent = userList.UserContent ? userList.UserContent : [];
		var listStat = userList.ListStat ? userList.ListStat : {};
		var errorMsg = userList.ErrorMsg ? userList.ErrorMsg : '';
		$.each(userContent, function (index, value) {
			if(value.Comments !== undefined){
				trStr += '<tr><td >' + value.Email + '</td><td>' + value.ReName + '</td><td>' + value.ListTitle + '</td><td>' + value.Comments + '</td><td>' + value.Type + '</td></tr>';
			}else{
				trStr += '<tr><td >' + value.Email + '</td><td>' + value.ReName + '</td><td>' + value.ListTitle + '</td><td>' + value.Type + '</td></tr>'
			}
		});
		$('.loading').hide();
		if(trStr){
			$('#list').html('');
			$('#list').html(trStr);
		}
		if(listStat) {
			$('#Comments').html(listStat.Comments);
			$('#Reviewers').html(listStat.Reviewers);
		}
		if(errorMsg){
			$('#ErrorMsg').html(errorMsg);
		}
	}catch (ex){
		alert(ex.message);
	}
}

function getHtmlView(){
	var htmlView = '<!DOCTYPE html><html lang="en">'+$('html').html()+'</html>';
	htmlView = htmlView.replace(/<script.*?>.*?<\/script>/ig, '');
	htmlView = htmlView.replace(/src="([^"]+)\.([^"]+)"/g, 'src="FoxTracker/$1.$2"');
	htmlView = htmlView.replace(/href="([^"]+)\.css"/g, 'href="FoxTracker/$1.css"');
	return htmlView;
}