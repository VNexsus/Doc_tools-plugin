/*
 (c) VNexsus 2021-2022

 Licensed under the Apache License, Version 2.0 (the "License");
 you may not use this file except in compliance with the License.
 You may obtain a copy of the License at

     http://www.apache.org/licenses/LICENSE-2.0

 Unless required by applicable law or agreed to in writing, software
 distributed under the License is distributed on an "AS IS" BASIS,
 WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 See the License for the specific language governing permissions and
 limitations under the License.
 */
 
(function(window, undefined){

	var text = '';

	function updateScroll()
	{
		Ps.update();
	}

	var _tools = [
		[ "Очистить колонтитулы", "clear_titles", 64, 64, true, false, false ],
		[ "Раскладка En>Ru", "layout_en_ru", 64, 64, true, true, true ],
		[ "Раскладка Ru>En", "layout_ru_en", 64, 64, true, true, true ],
		[ "Лишние пробелы", "extra_spaces", 64, 64, true, false, false ],
		[ "Сумма прописью", "cur2words", 64, 64, true, true, true ],
		[ "Число прописью", "num2words", 64, 64, true, true, true ]
	];

	function add_tools()
	{
		var _width = 0;
		for (var i = 0; i < _tools.length; i++)
		{
			if (_tools[i][2] > _width)
				_width = _tools[i][2];
		}

		_width += 20;

		var _space = 20;
		var _naturalWidth = window.innerWidth;

		var _count = ((_naturalWidth - _space) / (_width + _space)) >> 0;
		if (_count < 1)
			_count = 1;

		var _countRows = ((_tools.length + (_count - 1)) / _count) >> 0;

		var _html = "";
		var _index = 0;

		var _margin = (_naturalWidth - _count * (_width + _space)) >> 1;
		document.getElementById("main").style.marginLeft = _margin + "px";

		for (var _row = 0; _row < _countRows && _index < _tools.length; _row++)
		{
			_html += "<tr style='margin-left: " + _margin + "'>";

			for (var j = 0; j < _count; j++)
			{
				var _cur = _tools[_index];

				_html += "<td width='" + _width + "' height='" +_width + "' style='margin:" + (_space >> 1) + "'>";

				var _w = _cur[1];
				var _h = _cur[2];
				
				var disabled = ""; 
				
				switch (window.Asc.plugin.info.editorType) {
					case "word":
						disabled = _cur[4] ? "" : " class=\"disabled\"";
						break;
					case "cell":
						disabled = _cur[5] ? "" : " class=\"disabled\"";
						break;
					case "slide":
						disabled = _cur[6] ? "" : " class=\"disabled\"";
						break;
				}

				_html += ("<img id='tool" + _index + "' src=\"./tools/" + _cur[1] + "/icon.png\""+ disabled +"/>");
				_html += ("<div class=\"noselect celllabel\">" + _cur[0] + "</div>");

				_html += "</td>";

				_index++;

				if (_index >= _tools.length)
					break;
			}

			_html += "</tr>";
		}

		document.getElementById("main").innerHTML = _html;

		for (_index = 0; _index < _tools.length; _index++)
		{
			switch (window.Asc.plugin.info.editorType) {
				case "word":
					if(_tools[_index][4])
						document.getElementById("tool" + _index).onclick = new Function("return window."+ _tools[_index][1] +"();");
					break
				case "cell":
					if(_tools[_index][5])
						document.getElementById("tool" + _index).onclick = new Function("return window."+ _tools[_index][1] +"();");
					break;
				case "slide":
					if(_tools[_index][6])
						document.getElementById("tool" + _index).onclick = new Function("return window."+ _tools[_index][1] +"();");
					break;
			}
		}

		updateScroll();
	}
	
	window.onresize = function()
	{
		add_tools();
	};

	window.Asc.plugin.init = function(sText)
	{
		var container = document.getElementById('scrollable-container-id');
		
		Ps = new PerfectScrollbar('#' + container.id, {});
		
		add_tools();
		
		text = sText;
	};
	
	window.Asc.plugin.button = function(id)
	{
		this.executeCommand("close", "");
	};


	window.clear_titles = function(){
		
		window.Asc.plugin.info.recalculate = true;
		window.Asc.plugin.callCommand(function(){
			var oDocument = Api.GetDocument();
			var oSections = oDocument.GetSections();
			for(var i = 0; i < oSections.length; i++){
				oSections[i].RemoveHeader("default");
				oSections[i].RemoveHeader("title");
				oSections[i].RemoveHeader("even");
				oSections[i].RemoveFooter("default");
				oSections[i].RemoveFooter("title");
				oSections[i].RemoveFooter("even");
			}
		}, false, true, function(){});
		window.parent.focus();
	}


	window.layout_en_ru = function(){
		switch (window.Asc.plugin.info.editorType) {
			case "word":
				window.Asc.plugin.executeMethod("GetSelectedText", [], function(sText) {
					var en_chars = "`qwertyuiop[]asdfghjkl;'zxcvbnm,.~QWERTYUIOP{}ASDFGHJKL:\"|ZXCVBNM<>?"
					var ru_chars = "ёйцукенгшщзхъфывапролджэячсмитьбюЁЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭ/ЯЧСМИТЬБЮ,"
					var selected_text = "";
					var corrected_text = "";
					selected_text = sText;
					for(var i = 0; i < selected_text.length; i++){
						var _char = selected_text.charAt(i);
						var _index = en_chars.indexOf(_char);
						if(_index > 0)
							_char = ru_chars.charAt(_index);
						corrected_text = corrected_text + _char;
					}
					window.Asc.plugin.executeMethod("PasteText", [corrected_text]);
				});
				break;
			case "cell":
				var en_chars = "`qwertyuiop[]asdfghjkl;'zxcvbnm,.~QWERTYUIOP{}ASDFGHJKL:\"|ZXCVBNM<>?"
				var ru_chars = "ёйцукенгшщзхъфывапролджэячсмитьбюЁЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭ/ЯЧСМИТЬБЮ,"
				var oWorksheet = parent.g_asc_plugins.api.GetActiveSheet();
				if(oWorksheet){
					var aRange = oWorksheet.GetSelection().GetAddress(false,false,"xlA1",false);
					var aRangeAddr = aRange.split(':');
					var colstart = oWorksheet.GetRange(aRangeAddr[0]).GetCol();
					var rowstart = oWorksheet.GetRange(aRangeAddr[0]).GetRow();
					var colend = colstart;
					var rowend = rowstart;
					if(aRangeAddr.length > 1){
						var colend = oWorksheet.GetRange(aRangeAddr[1]).GetCol();
						var rowend = oWorksheet.GetRange(aRangeAddr[1]).GetRow();
					}
					for(var x = colstart; x <= colend; x++)
						for(var y = rowstart; y <= rowend; y++){
							var cell = oWorksheet.GetRangeByNumber(y,x);
							var selected_text = cell.GetText();
							if(selected_text){
								var corrected_text = "";
								for(var i = 0; i < selected_text.length; i++){
									var _char = selected_text.charAt(i);
									var _index = en_chars.indexOf(_char);
									if(_index > 0)
										_char = ru_chars.charAt(_index);
									corrected_text = corrected_text + _char;
								}
								cell.SetValue(corrected_text);
							}

						}
					parent.Asc.editor.controller.view.resize();
				}
				break;
			case "slide":
				window.Asc.plugin.executeMethod("GetSelectedText", [], function(sText) {
					var en_chars = "`qwertyuiop[]asdfghjkl;'zxcvbnm,.~QWERTYUIOP{}ASDFGHJKL:\"|ZXCVBNM<>?"
					var ru_chars = "ёйцукенгшщзхъфывапролджэячсмитьбюЁЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭ/ЯЧСМИТЬБЮ,"
					var selected_text = "";
					var corrected_text = "";
					selected_text = text;
					for(var i = 0; i < selected_text.length; i++){
						var _char = selected_text.charAt(i);
						var _index = en_chars.indexOf(_char);
						if(_index > 0)
							_char = ru_chars.charAt(_index);
						corrected_text = corrected_text + _char;
					}
					window.Asc.plugin.executeMethod("PasteText", [corrected_text]);
				});
				break;
		}
		window.parent.focus();
	}


	window.layout_ru_en = function(){
		switch (window.Asc.plugin.info.editorType) {
			case "word":
				window.Asc.plugin.executeMethod("GetSelectedText", [], function(sText) {
					var en_chars = "`qwertyuiop[]asdfghjkl;'zxcvbnm,.~QWERTYUIOP{}ASDFGHJKL:\"|ZXCVBNM<>?"
					var ru_chars = "ёйцукенгшщзхъфывапролджэячсмитьбюЁЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭ/ЯЧСМИТЬБЮ,"
					var selected_text = "";
					var corrected_text = "";
					selected_text = sText;
					for(var i = 0; i < selected_text.length; i++){
						var _char = selected_text.charAt(i);
						var _index = ru_chars.indexOf(_char);
						if(_index > 0)
							_char = en_chars.charAt(_index);
						corrected_text = corrected_text + _char;
					}
					window.Asc.plugin.executeMethod("PasteText", [corrected_text]);
				});
				break;
			case "cell":
				var en_chars = "`qwertyuiop[]asdfghjkl;'zxcvbnm,.~QWERTYUIOP{}ASDFGHJKL:\"|ZXCVBNM<>?"
				var ru_chars = "ёйцукенгшщзхъфывапролджэячсмитьбюЁЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭ/ЯЧСМИТЬБЮ,"
				var oWorksheet = parent.g_asc_plugins.api.GetActiveSheet();
				if(oWorksheet){
					var aRange = oWorksheet.GetSelection().GetAddress(false,false,"xlA1",false);
					var aRangeAddr = aRange.split(':');
					var colstart = oWorksheet.GetRange(aRangeAddr[0]).GetCol();
					var rowstart = oWorksheet.GetRange(aRangeAddr[0]).GetRow();
					var colend = colstart;
					var rowend = rowstart;
					if(aRangeAddr.length > 1){
						var colend = oWorksheet.GetRange(aRangeAddr[1]).GetCol();
						var rowend = oWorksheet.GetRange(aRangeAddr[1]).GetRow();
					}
					for(var x = colstart; x <= colend; x++)
						for(var y = rowstart; y <= rowend; y++){
							var cell = oWorksheet.GetRangeByNumber(y,x);
							var selected_text = cell.GetText();
							if(selected_text){
								var corrected_text = "";
								for(var i = 0; i < selected_text.length; i++){
									var _char = selected_text.charAt(i);
									var _index = ru_chars.indexOf(_char);
									if(_index > 0)
										_char = en_chars.charAt(_index);
									corrected_text = corrected_text + _char;
								}
								cell.SetValue(corrected_text);
							}

						}
					parent.Asc.editor.controller.view.resize();
				}
				break;
			case "slide":
				window.Asc.plugin.executeMethod("GetSelectedText", [], function(sText) {
					var en_chars = "`qwertyuiop[]asdfghjkl;'zxcvbnm,.~QWERTYUIOP{}ASDFGHJKL:\"|ZXCVBNM<>?"
					var ru_chars = "ёйцукенгшщзхъфывапролджэячсмитьбюЁЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭ/ЯЧСМИТЬБЮ,"
					var selected_text = "";
					var corrected_text = "";
					selected_text = text;
					for(var i = 0; i < selected_text.length; i++){
						var _char = selected_text.charAt(i);
						var _index = ru_chars.indexOf(_char);
						if(_index > 0)
							_char = en_chars.charAt(_index);
						corrected_text = corrected_text + _char;
					}
					window.Asc.plugin.executeMethod("PasteText", [corrected_text]);
				});
				break;
		}
		window.parent.focus();
	}


	window.extra_spaces = function(){
		var oProperties = {
			"searchString"  : "     ",
			"replaceString" : " ",
			"matchCase"     : false
		};
		window.Asc.plugin.info.recalculate = true;
		window.Asc.plugin.executeMethod("SearchAndReplace", [oProperties]);
		var oProperties = {
			"searchString"  : "    ",
			"replaceString" : " ",
			"matchCase"     : false
		};
		window.Asc.plugin.info.recalculate = true;
		window.Asc.plugin.executeMethod("SearchAndReplace", [oProperties]);
		var oProperties = {
			"searchString"  : "   ",
			"replaceString" : " ",
			"matchCase"     : false
		};
		window.Asc.plugin.info.recalculate = true;
		window.Asc.plugin.executeMethod("SearchAndReplace", [oProperties]);
		var oProperties = {
			"searchString"  : "  ",
			"replaceString" : " ",
			"matchCase"     : false
		};
		window.Asc.plugin.info.recalculate = true;
		window.Asc.plugin.executeMethod("SearchAndReplace", [oProperties]);
		var oProperties = {
			"searchString"  : " ,",
			"replaceString" : ",",
			"matchCase"     : false
		};
		window.Asc.plugin.info.recalculate = true;
		window.Asc.plugin.executeMethod("SearchAndReplace", [oProperties]);
		var oProperties = {
			"searchString"  : " .",
			"replaceString" : ",",
			"matchCase"     : false
		};
		window.Asc.plugin.info.recalculate = true;
		window.Asc.plugin.executeMethod("SearchAndReplace", [oProperties]);
		var oProperties = {
			"searchString"  : "( ",
			"replaceString" : "(",
			"matchCase"     : false
		};
		window.Asc.plugin.info.recalculate = true;
		window.Asc.plugin.executeMethod("SearchAndReplace", [oProperties]);
		var oProperties = {
			"searchString"  : " )",
			"replaceString" : ")",
			"matchCase"     : false
		};
		window.Asc.plugin.info.recalculate = true;
		window.Asc.plugin.executeMethod("SearchAndReplace", [oProperties]);
		var patterns = [" в", " без", " до", " из", " к", " на", " по", " о", " от", " перед", " при", " через", " с", " у", " за", " над", " об", " под", " про", " для", "№", "АО", "ООО"]
		patterns.forEach(function(pattern){
			var oProperties = {
				"searchString"  : pattern + " ",
				"replaceString" : pattern + " ",
				"matchCase"     : true
			};
			window.Asc.plugin.info.recalculate = true;
			window.Asc.plugin.executeMethod("SearchAndReplace", [oProperties]);
		});
		window.parent.focus();
	}

	window.cur2words = function(){
		switch (window.Asc.plugin.info.editorType) {
			case "word":
				window.Asc.plugin.executeMethod("GetSelectedText", [], function(sText) {
					var converted_text = window.convert(sText);
					if(converted_text != "")
						window.Asc.plugin.executeMethod("PasteText", [sText + " " + converted_text]);
				});
				break;
			case "cell":
				var oWorksheet = parent.g_asc_plugins.api.GetActiveSheet();
				if(oWorksheet){
					var aRange = oWorksheet.GetSelection().GetAddress(false,false,"xlA1",false);
					var aRangeAddr = aRange.split(':');
					var colstart = oWorksheet.GetRange(aRangeAddr[0]).GetCol();
					var rowstart = oWorksheet.GetRange(aRangeAddr[0]).GetRow();
					if(parent.Asc.editor.GetVersion() == "7.4.0"){
						colstart--;
						rowstart--;
					}
					var colend = colstart;
					var rowend = rowstart;
					if(aRangeAddr.length > 1){
						var colend = oWorksheet.GetRange(aRangeAddr[1]).GetCol();
						var rowend = oWorksheet.GetRange(aRangeAddr[1]).GetRow();
						if(parent.Asc.editor.GetVersion() == "7.4.0"){
							colend--;
							rowend--;
						}
						if(colend > colstart){
							parent.Common.UI.warning({msg: "Выделите только один столбец!"})
							return;
						}
					}
					for(var x = colstart; x <= colend; x++)
						for(var y = rowstart; y <= rowend; y++){
							var cell = oWorksheet.GetRangeByNumber(y,x);
							var text = cell.Text;
							if(text){
								var matched = text.match(/(\D{0,1})([\d\.\, ]+)\s*(\S*)/);
								if(matched && matched.length > 0){
									switch(matched[1] + matched[3]){
										case "$":
										case "USD":
										case " USD":
											oWorksheet.GetRangeByNumber(y, x+1).SetValue(window.convert(matched[2],{currency: 'usd'}));
											break;
										case " €":
										case "€":
										case "EUR":
										case " EUR":
											oWorksheet.GetRangeByNumber(y, x+1).SetValue(window.convert(matched[2],{currency: 'eur'}));
											break;
										case "¥":
										case "元":
										case "圆":
										case "圓":
										case "CNY":
										case " CNY":
											oWorksheet.GetRangeByNumber(y, x+1).SetValue(window.convert(matched[2],{currency: {currencyNameCases: ['юань', 'юаня', 'юаней'], fractionalPartNameCases: ['фынь', 'фыня', 'фыней'], currencyNounGender: {integer: 0, fractionalPart: 0}}}));
											break;
										default:
											oWorksheet.GetRangeByNumber(y, x+1).SetValue(window.convert(matched[2]));
											break;
									}
								}
							}

						}
				parent.Asc.editor.controller.view.resize();
				}
				break;
			case "slide":
				window.Asc.plugin.executeMethod("GetSelectedText", [], function(sText) {
					var converted_text = window.convert(sText);
					if(converted_text != "")
						window.Asc.plugin.executeMethod("PasteText", [sText + " " + converted_text]);
				});
				break;
		}
		window.parent.focus();
	}

	window.num2words = function(){
		switch (window.Asc.plugin.info.editorType) {
			case "word":
				window.Asc.plugin.executeMethod("GetSelectedText", [], function(sText) {
					var converted_text = window.convert(sText, {currency: 'number'});
					if(converted_text != "")
						window.Asc.plugin.executeMethod("PasteText", [sText + " " + converted_text]);
				});
				break;
			case "cell":
				var oWorksheet = parent.g_asc_plugins.api.GetActiveSheet();
				if(oWorksheet){
					var aRange = oWorksheet.GetSelection().GetAddress(false,false,"xlA1",false);
					var aRangeAddr = aRange.split(':');
					var colstart = oWorksheet.GetRange(aRangeAddr[0]).GetCol();
					var rowstart = oWorksheet.GetRange(aRangeAddr[0]).GetRow();
					if(parent.Asc.editor.GetVersion() == "7.4.0"){
						colstart--;
						rowstart--;
					}
					var colend = colstart;
					var rowend = rowstart;
					if(aRangeAddr.length > 1){
						var colend = oWorksheet.GetRange(aRangeAddr[1]).GetCol();
						var rowend = oWorksheet.GetRange(aRangeAddr[1]).GetRow();
						if(parent.Asc.editor.GetVersion() == "7.4.0"){
							colend--;
							rowend--;
						}
						if(colend > colstart){
							parent.Common.UI.warning({msg: "Выделите только один столбец!"})
							return;
						}
					}
					for(var x = colstart; x <= colend; x++)
						for(var y = rowstart; y <= rowend; y++){
							var cell = oWorksheet.GetRangeByNumber(y,x);
							var text = cell.GetText();
							if(text)
								oWorksheet.GetRangeByNumber(y, x+1).SetValue(window.convert(text, {currency: 'number'}));
						}
					parent.Asc.editor.controller.view.resize();
				}
				break;
			case "slide":
				window.Asc.plugin.executeMethod("GetSelectedText", [], function(sText) {
					var converted_text = window.convert(sText, {currency: 'number'});
					if(converted_text != "")
						window.Asc.plugin.executeMethod("PasteText", [sText + " " + converted_text]);
				});
				break;
		}
		window.parent.focus();
	}

	window.Asc.plugin.onExternalMouseUp = function()
    {
        var evt = document.createEvent("MouseEvents");
        evt.initMouseEvent("mouseup", true, true, window, 1, 0, 0, 0, 0,
            false, false, false, false, 0, null);

        document.dispatchEvent(evt);
    };

})(window, undefined);
