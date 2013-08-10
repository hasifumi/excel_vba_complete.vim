function! neocomplcache#sources#test#define()"{{{
        return s:source
endfunction"}}}

let s:source = {
  \ 'name' : 'test',
  \ 'kind' : 'ftplugin',
  \ 'filetypes' : { 'vb': 1, 'basic': 1  }
  \ } "DON'T FOLDING! it will occure an error.

function! s:source.initialize()"{{{
  let s:keywords = []
endfunction"}}}

function! s:source.finalize()"{{{
  unlet s:keywords
endfunction"}}}

function! s:source.get_keyword_pos(cur_text)"{{{
  "if neocomplcache#within_comment()
  "  return -1
  "endif
  call add(s:keywords, {'word': 'test', 'menu': 'test'})
  return s:ret
endfunction"}}}

function! s:source.get_complete_words(cur_keyword_pos, cur_keyword_str)"{{{
  return neocomplcache#keyword_filter(copy(s:keywords), a:cur_keyword_str)
endfunction"}}}

