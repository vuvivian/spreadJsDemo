window.cubeReportExcelImportFile = function (spread) {

$(document.body).append("<form id= \"uploadForm_excel\"> <input type=\"file\" style=\"display:none;\" id=\"file_id_excel\" name=\"file\"  accept=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\"/> </form> ");

	$("#file_id_excel").on('change',function(e){ //btn_file为隐藏的input
		if ($(this).val() == '') { //如果没有选择文件则不触发
			return false;
		} else {
			var sheet = spread.getActiveSheet();
			var excelIo = new GC.Spread.Excel.IO();
			var excelFile = document.getElementById("file_id_excel").files[0];
			excelIo.open(excelFile, function (json) {
				var workbookObj = json;
				spread.fromJSON(workbookObj);
			}, function (e) {
				alert(e.errorMessage);
			});
			$("#uploadForm_excel").remove();
		}
	});
document.getElementById("file_id_excel").click();
}