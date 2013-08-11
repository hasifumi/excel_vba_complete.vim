" vim:set foldmethod=marker :

function! neocomplcache#sources#excel_vba_complete#define()"{{{
        return s:source
endfunction"}}}

let s:source = {
  \ 'name' : 'excel_vba_complete',
  \ 'kind' : 'ftplugin',
  \ 'filetypes' : { 'vb': 1, 'basic': 1  }
  \ } "DON'T FOLDING! it will occure an error.

function! s:source.initialize()"{{{
  let s:keywords = []
  let s:objects = {
 \  'Workbook': { 
 \    'create': 'Dim',
 \    'property': {
 \       'Name': {
 \         'kind': 'v',
 \         'info': '',
 \       },
 \       'Path': {
 \         'kind': 'v',
 \         'info': '',
 \       },
 \       'Worksheets': {
 \         'kind': 'o',
 \         'info': '',
 \         'name': 'Worksheets'
 \       },
 \       'ActiveSheet': {
 \         'kind': 'o',
 \         'info': '',
 \         'name': 'Worksheets'
 \       },
 \    },
 \    'method': {
 \       'Delete': {
 \         'kind': 'f',
 \         'info': '',
 \       },
 \       'Save': {
 \         'kind': 'f',
 \         'info': '',
 \       },
 \       'SaveAs': {
 \         'kind': 'f',
 \         'info': '',
 \       },
 \       'Close': {
 \         'kind': 'f',
 \         'info': '',
 \       },
 \    },
 \  },
 \  'Workbooks': { 
 \    'create': 'Dim',
 \    'property': {
 \       'Count': {
 \         'kind': 'v',
 \         'info': '',
 \       },
 \    },
 \    'method': {
 \       'Add': {
 \         'kind': 'f',
 \         'info': '',
 \       },
 \       'Open': {
 \         'kind': 'f',
 \         'info': '',
 \       },
 \    },
 \  },
 \  'Worksheet': { 
 \    'create': 'Dim',
 \    'property': {
 \       'Name': {
 \         'kind': 'v',
 \         'info': '',
 \       },
 \       'Range': {
 \         'kind': 'o',
 \         'info': '',
 \         'name': 'Range'
 \       },
 \       'Cells': {
 \         'kind': 'o',
 \         'info': '',
 \         'name': 'Range'
 \       },
 \       'ActiveCell': {
 \         'kind': 'o',
 \         'info': '',
 \         'name': 'Range'
 \       },
 \       'Rows': {
 \         'kind': 'v',
 \         'info': '',
 \         'property': ['Count', 'Height', 'AutoFit()'],
 \       },
 \       'Columns': {
 \         'kind': 'v',
 \         'info': '',
 \         'property': ['Count', 'Width', 'AutoFit()'],
 \       },
 \    },
 \    'method': {
 \       'Paste': {
 \         'kind': 'f',
 \         'info': '',
 \       },
 \       'Delete': {
 \         'kind': 'f',
 \         'info': '',
 \       },
 \       'Move': {
 \         'kind': 'f',
 \         'info': '',
 \       },
 \       'Copy': {
 \         'kind': 'f',
 \         'info': '',
 \       },
 \    },
 \  },
 \  'Worksheets': { 
 \    'create': 'Dim',
 \    'property': {
 \       'Count': {
 \         'kind': 'v',
 \         'info': '',
 \       },
 \    },
 \    'method': {
 \       'Add': {
 \         'kind': 'f',
 \         'info': '',
 \       },
 \    },
 \  },
 \  'Range': { 
 \    'create': 'Dim',
 \    'property': {
 \       'Value': {
 \         'kind': 'v',
 \         'info': '',
 \       },
 \       'End': {
 \         'kind': 'v',
 \         'info': '',
 \         'args': ['xlUp', 'xlDown', 'xltoRight', 'xlToLeft'],
 \       },
 \       'Rows': {
 \         'kind': 'o',
 \         'info': '',
 \         'property': ['Count', 'Height', 'AutoFit()']
 \       },
 \       'Columns': {
 \         'kind': 'o',
 \         'info': '',
 \         'property': ['Count', 'Width', 'AutoFit()']
 \       },
 \    },
 \    'method': {
 \       'Clear': {
 \         'kind': 'f',
 \         'info': '',
 \       },
 \    },
 \  },
 \}
"  let s:objects = {"{{{
" \  'Workbook': { "{{{
" \    'create': 'Dim',
" \    'member': { 'Name' : 'v', 
" \                'Path' : 'v' , 
" \                'Worksheets' : 'v' , 
" \                'ActiveSheet' : 'v' , 
" \                'Delete' : 'f' , 
" \                'SaveAs' : 'f' , 
" \                'Save' : 'f' , 
" \                'Close' : 'f' , 
" \    },
" \  },"}}}
" \  'Workbooks': { "{{{
" \    'create': '',
" \    'member': { 'Count' : 'v', 
" \                'Add' : 'f' , 
" \                'Open' : 'f' , 
" \    },
" \  },"}}}
" \  'Worksheet': { "{{{
" \    'create': 'Dim',
" \    'member': { 'Name' : 'v', 
" \                'Range' : 'o' , 
" \                'Cells' : 'o' , 
" \                'Rows' : 'o' , 
" \                'Columns' : 'o' , 
" \                'ActiveCell' : 'o' , 
" \                'Paste' : 'f' , 
" \                'Delete' : 'f' , 
" \                'Move' : 'f' , 
" \                'Copy' : 'f' , 
" \    },
" \  },"}}}
" \  'Worksheets': { "{{{
" \    'create': 'Dim',
" \    'member': { 'Counts' : 'v', 
" \                'Add' : 'f' , 
" \    },
" \  },"}}}
" \  'Range': { 
" \    'create': 'Dim',
" \    'member': { 'Value' : 'v', 
" \                'Rows' : 'v' , 
" \                'Columns' : 'v' , 
" \                'EntireRow' : 'v' , 
" \                'EntireColumn' : 'v' , 
" \                'Row' : 'v' , 
" \                'Column' : 'v' , 
" \                'Count' : 'v' , 
" \                'Offset' : 'v' , 
" \                'CurrentRegion' : 'v' , 
" \                'End' : 'v' , 
" \                'RowHeight' : 'v' , 
" \                'ColumnWidth' : 'v' , 
" \                'Formula' : 'v' , 
" \                'FormulaR1C1' : 'v' , 
" \                'Font' : 'o' , 
" \                'Interior' : 'o' , 
" \                'Select' : 'f' , 
" \                'Activate' : 'f' , 
" \                'Cut' : 'f' , 
" \                'Copy' : 'f' , 
" \                'PasteSpecial' : 'f' , 
" \                'Insert' : 'f' , 
" \                'Delete' : 'f' , 
" \                'AutoFill' : 'f' , 
" \                'Clear' : 'f' , 
" \                'ClearContents' : 'f' , 
" \                'ClearFormats' : 'f' , 
" \                'Sort' : 'f' , 
" \                'AutoFilter' : 'f' , 
" \    },
" \  },
" \  'Application': { 
" \    'create': '',
" \    'member': { 'Selection' : 'v', 
" \                'ActiveCell' : 'v' , 
" \                'CutCopyMode' : 'v' , 
" \                'Statusbar' : 'v' , 
" \                'DisplayAlerts' : 'v' , 
" \                'WorksheetFunction' : 'o' , 
" \    },
" \  },
" \  'Rows': { 
" \    'create': '',
" \    'member': { 'Height' : 'f', 
" \                'AutoFit' : 'v' , 
" \    },
" \  },
" \  'Columns': { 
" \    'create': '',
" \    'member': { 'Width' : 'f', 
" \                'AutoFit' : 'v' , 
" \    },
" \  },
" \  'Font': { 
" \    'create': '',
" \    'member': { 'Name' : 'v', 
" \                'Size' : 'v' , 
" \                'Bold' : 'v' , 
" \                'Italic' : 'v' , 
" \                'Underline' : 'v' , 
" \                'ColorIndex' : 'v' , 
" \    },
" \  },
" \  'Interior': { 
" \    'create': '',
" \    'member': { 'Color' : 'v', 
" \                'ColorIndex' : 'v' , 
" \    },
" \  },
" \}"}}}
  let s:variables = {}
  let s:line = 0
  let s:temp_objects = {}
endfunction"}}}

function! s:source.finalize()"{{{
  unlet s:objects
  unlet s:temp_objects
  unlet s:keywords
  unlet s:line
  unlet s:variables
endfunction"}}}

function! excel_vba_complete#initialize()"{{{
  call s:source.initialize()
endfunction"}}}

function! s:source.get_keyword_pos(cur_text)"{{{
  if neocomplcache#within_comment()
    return -1
  endif
  if &modified
    call excel_vba_complete#get_all_variables()
  endif
  for word1 in keys(s:variables)
    if a:cur_text == word1
      call excel_vba_complete#gather_keywords(s:keywords, word1, 'property')
      call excel_vba_complete#gather_keywords(s:keywords, word1, 'method')
      "for word in keys(s:objects[s:variables[word1]['type']]['property'])"{{{
      "  call add(s:keywords, { 
      "   \ 'word' : word1.'.'.word,
      "   \ 'abbr': word, 
      "   \ 'menu': '[excel_vba_complete]', 
      "   \ 'kind' : s:objects[s:variables[word1]['type']]['property'][word]['kind']
      "   \ })
      "endfor
      "for word in keys(s:objects[s:variables[word1]['type']]['method'])
      "  call add(s:keywords, { 
      "   \ 'word' : word1.'.'.word,
      "   \ 'abbr': word, 
      "   \ 'menu': '[excel_vba_complete]', 
      "   \ 'kind' : s:objects[s:variables[word1]['type']]['method'][word]['kind']
      "   \ })
      "endfor"}}}
      return match(a:cur_text, word1.".")
      break
    endif
  endfor
  "for word1 in keys(s:variables)
  "  if a:cur_text =~ word1
  "    for word in keys(s:objects[s:variables[word1]['type']]['member'])
  "      "echo "add " . word1 . "." . word . " to s:keywords"
  "      call add(s:keywords, { 'word' : word1.".".word,
  "       \ 'abbr': word, 
  "       \ 'menu': '[excel_vba_complete]', 
  "       \ 'kind' : s:objects[s:variables[word1]['type']]['member'][word]})
  "    endfor
  "    return match(a:cur_text, word1.".")
  "    break
  "  endif
  "endfor
endfunction"}}}

function! s:source.get_complete_words(cur_keyword_pos, cur_keyword_str)"{{{
  return neocomplcache#keyword_filter(copy(s:keywords), a:cur_keyword_str)
endfunction"}}}

function! excel_vba_complete#gather_keywords(dict, word, flg)"{{{
  for key in keys(s:objects[s:variables[a:word]['type']][a:flg])
    call add(a:dict, { 
     \ 'word' : a:word . '.' . key,
     \ 'abbr': key, 
     \ 'menu': '[excel_vba_complete]', 
     \ 'kind' : s:objects[s:variables[a:word]['type']][a:flg][key]['kind']
     \ })
  endfor
endfunction"}}}

function! excel_vba_complete#get_variables(line)"{{{
  "let temp_line = substitute(a:line, '\s', '', 'g')
  "echo temp_line
  "if a:line=~ 'Dim' || a:line =~ 'dim'
  "  let list = matchlist(a:line, '\s*[Dim|dim]\s*\(\w*\)\s*[As|as]\s*\(\w*\)')
  if a:line=~ 'Dim'
    "echo a:line
    let list = matchlist(a:line, '\s*Dim\s*\(\w*\)\s*As\s*\(\w*\)')
    "echo list
    for k in keys(s:objects)
      "echo k
      if (len(list) > 0) && (k =~ list[2])
        if !has_key(s:variables, list[1])
          let s:variables[list[1]] = { 'type': list[2] }
          "echo "s:variables[list[1]][type]:" . s:variables[list[1]]['type']
        endif  
      endif  
    endfor
  endif
endfunction"}}}

function! excel_vba_complete#get_all_variables()"{{{
  let s:variables = {}
  let lines = getline(0, line("$"))
  "for line in lines
  "  call excel_vba_complete#find_require_line(line)
  "endfor
  for line in lines
    call excel_vba_complete#get_variables(line)
  endfor
endfunction"}}}

function! excel_vba_complete#show_all_variables()"{{{
  for i in keys(s:variables)
    echo ' key: ' . i . ', type: ' . s:variables[i]['type']
  endfor
endfunction"}}}

function! excel_vba_complete#show_all_objects()"{{{
  for i in keys(s:objects)
    echo ' key: ' . i 
  endfor
endfunction"}}}

function! excel_vba_complete#show_objects(object)"{{{
  echo s:objects[a:object]
endfunction"}}}

function! excel_vba_complete#show_all_temp_objects()"{{{
  for i in keys(s:temp_objects)
    echo s:temp_objects[i]
  endfor 
endfunction"}}}

function! excel_vba_complete#show_temp_object(object)"{{{
  echo s:temp_objects[a:object]
endfunction"}}}

function! excel_vba_complete#test(word)"{{{"{{{
  echo s:variables[a:word]['type']
  echo s:temp_objects[s:variables[a:word]['type']]['member']
endfunction"}}}"}}}

function! excel_vba_complete#add_temp_object(class, member, kind)"{{{
  "echo "class:" . a:class . ", member:" . a:member . ", kind:" . a:kind
  if has_key(s:objects, a:class)
    "echo "has class"
    if has_key(s:objects[a:class]["member"], a:member)
      "echo "has member"
    else
      let s:objects[a:class]["member"][a:member] = a:kind
    endif
  else
    if empty(a:member) || empty(a:kind)
      let s:objects[a:class] = {'member':{}, 'create':'new'}
    else
      let s:objects[a:class] = {'member':{}, 'create':'new'}
      let s:objects[a:class]["member"][a:member] = a:kind
    endif
  endif
endfunction"}}}

"function! excel_vba_complete#find_require_line(line)"{{{
"  let aft0 = substitute(a:line, " ", "", "g")
"  "echo "aft0:" . aft0
"  let aft1 = substitute(aft0, "'", "\"", "g")
"  "echo "aft1:" . aft1
"  if aft1 =~ "require"
"    echo "found require"
"    let list = matchlist(aft1, '\(\w*\)=\w*("\(\w*\)"')
"    "echo list[2]
"    call excel_vba_complete#glob_require_file(list[2])
"  endif
"endfunction"}}}
"
"function! excel_vba_complete#glob_require_file(filename)"{{{
"  let base = "./" . a:filename . ".coffee"
"  "echo base
"  let filelist = glob(base)
"  let splitted = split(filelist)
"  for file in splitted
"    "echo file
"    if filereadable(file)
"      echo "readable!"
"      for line in readfile(file)
"        "echo line
"        let res = excel_vba_complete#find_member_line(line)
"        "echo res
"        if !empty(res)
"          "echo 'res[0]:' . res[0] . ', res[1]:' . res[1]
"          call excel_vba_complete#add_temp_object("temp_" . a:filename, res[0], res[1])
"        endif
"      endfor
"    endif
"  endfor
"endfunction"}}}
"
"function! excel_vba_complete#find_member_line(line)"{{{
"  let res = []
"  let aft0 = substitute(a:line, " ", "", "g")
"  if aft0 =~ "->" || aft0 =~ "=>"
"    echo "found coffee function"
"    let list = matchlist(aft0, '\w*\.\(\w*\)=\.*')
"    let res = [list[1],'f']
"  "elseif aft0 =~ "=" && aft0 =~ "self\."
"  "  echo "found coffee property"
"  "  let list = matchlist(aft0, '\w*\.\(\w*\)=\.*')
"  "  "echo list[1]
"  "  let res = [list[1],'v']
"  endif
"  return res
"endfunction"}}}
