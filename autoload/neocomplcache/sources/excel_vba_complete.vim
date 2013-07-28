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
 \    'member': { 'Name' : 'v', 
 \                'Path' : 'v' , 
 \                'testFunc' : 'f' , 
 \    },
 \  },
 \  'Worksheet': { 
 \    'create': 'Dim',
 \    'member': { 'Name' : 'v', 
 \                'testFunc' : 'f' , 
 \    },
 \  },
 \}
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
    call excel_vba_complete#get_all_variables
  endif
  for word1 in keys(s:variables)
    if a:cur_text =~ word1
      for word in keys(s:objects[s:variables[word1]['type']]['member'])
        echo "add " . word1 . "." . word . " to s:keywords"
        call add(s:keywords, { 'word' : word1.".".word, 'menu': '[excel_vba_complete]', 
         \ 'kind' : s:objects[s:variables[word1]['type']]['member'][word]})
      endfor
      return match(a:cur_text, word1.".")
      break
    endif
  endfor
endfunction"}}}

function! s:source.get_complete_words(cur_keyword_pos, cur_keyword_str)"{{{
  return neocomplcache#keyword_filter(copy(s:keywords), a:cur_keyword_str)
endfunction"}}}

function! excel_vba_complete#get_variables(line)"{{{
  let temp_line = substitute(a:line, '\s', '', 'g')
  "echo temp_line
  if temp_line =~ '=' && ( temp_line=~ 'Dim' || temp_line =~ 'dim')
    "let list = matchlist(temp_line, '\(\w*\)\(=\)new\(\w*\)')
    let list = matchlist(temp_line, '\s[Dim,dim]\s\(\w*\)\s[As,as]\s\(\w*\)')
    "echo list
    "for k in keys(s:objects)
    "  "echo k
    "  if (len(list) > 0) && (k =~ list[3])
    "    if !has_key(s:variables, list[1])
    "      let s:variables[list[1]] = { 'type': list[3] }
    "      "echo "s:variables[list[1]][type]:" . s:variables[list[1]]['type']
    "    endif  
    "  endif  
    "endfor
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
