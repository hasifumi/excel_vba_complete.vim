function! Test1()
  let s:ret = 0
  let s:flg = 0
  let line = getline(".")
  let s:start = col(".") - 1 
  let s:end = s:start
  echo "START!!, s:start:" . s:start . ", s:end:" . s:end

  while s:start >= 0
    if line[s:start] =~ '[\.:[:blank:]]' 
      "let s:start -= 1
      if line[s:start] =~ '[\.]'
        let s:flg = 1
      endif
      break
    endif
    "if line[s:start - 1] == '.'
    "  break
    "endif
    let s:start -= 1
  endwhile

  echo "s:flg 0:not found dot, 1:found dot"
  if s:flg
    echo "True:" . s:flg
    echo "01234567890123456789012345678901234567890123456789"
    echo line 
    echo "END!!!, s:start+1:" . (s:start + 1) . ", s:end:" . s:end . ", length:" . (s:end - s:start)
    echo strpart(line, (s:start + 1), (s:end - s:start))
    let s:ret = s:start + 1
    echo "s:ret:" . s:ret
  else
    echo "False:" . s:flg
    echo "s:ret:" . s:ret
  endif
  return s:ret
endfunction 

command! Test1 :call Test1()

function! Test2()
  let s:cnt = 0
  let line = getline('.')
  let s:start = col('.')
  let s:end = s:start

  while s:start >= 0
    if line[s:start] =~ ')'
      let s:cnt += 1
      let s:start -= 1
    endif
    if line[s:start] =~ '(' && s:cnt > 0
      let s:cnt -= 1
      let s:start -= 1
      if s:cnt == 0
        let s:end = s:start
      endif
    endif
    if s:cnt == 0 && ( line[s:start - 1] =~ '[:blank:]' || line[s:start - 1] =~ ':' || line[s:start - 1] =~ '.')
      "let s:start -= 1
      break
    endif
  endwhile
  "echo "s:start-1:" . (s:start - 1) . ", s:end-1:" . (s:end - 1) . ", length:". (s:end - s:start + 1) 
  echo "s:start:" . s:start . ", s:end:" . s:end . ", length:". (s:end - s:start + 1) 
        \ . ", strpart:" . strpart(line, (s:start - 1), (s:end - s:start + 1))
  return strpart(line, (s:start - 1), (s:end - s:start + 1))
endfunction

command! Test2 :call Test2()

function! Test3()
  let s:cnt = 0
  let line = getline('.')
  let s:start = col('.') - 1
  let s:end = s:start
  echo "START!!, s:start:" . s:start . ", s:end:" . s:end . ", s:cnt:" . s:cnt
  if line[s:start] =~ '\.'
    let s:start -= 1
  endif
  while s:start >= 0
    if line[s:start] =~ ')'
      let s:cnt += 1
      echo "Match ), s:start:" . s:start . ", s:end:" . s:end . ", s:cnt:" . s:cnt
      "break
    endif
    if line[s:start] =~ '(' && s:cnt > 0
      let s:cnt -= 1
      echo "Match (, s:start:" . s:start . ", s:end:" . s:end . ", s:cnt:" . s:cnt
      if s:cnt == 0
        let s:end = s:start - 1
      endif
      "break
    endif
    "if line[s:start] =~ '(' && s:cnt == 0
    "  let s:end = s:start
    "endif
    "if s:cnt > 0 && line[s:start] =~ '.'
    "if line[s:start] =~ "."
    "if line[s:start] =~ '[:punct:]' && s:cnt == 0
    "if (line[s:start] =~ '[:punct:]' || line[s:start] =~ '[:blank:]') && s:cnt == 0
    "if line[s:start] =~ '\.' && s:cnt == 0
    "if (line[s:start] =~ '\.' || line[s:start] =~ '[:blank:]') && s:cnt == 0
    if line[s:start] =~ '[\.:[:blank:]]' 
      if line[s:start] =~ '[\.]'
        let s:flg = 1
      endif
      if s:cnt == 0
        "let s:start -= 1
        echo "Match dot or colon or space or tab, s:start:" . s:start . ", s:end:" . s:end . ", s:cnt:" . s:cnt
        break
      endif
    endif
    let s:start -= 1
  endwhile
  echo "01234567890123456789012345678901234567890123456789"
  echo line 
  echo "END!!!, s:start+1:" . (s:start + 1) . ", s:end:" . s:end . ", length:" . (s:end - s:start) . ", s:cnt:" . s:cnt
  "echo "END!!!, s:start:" . s:start . ", s:end:" . s:end . ", length:" . (s:end - s:start) . ", s:cnt:" . s:cnt
  echo strpart(line, (s:start + 1), (s:end - s:start))
endfunction

command! Test3 :call Test3()
