
function bossdie(x , min , max, iguess  )
bossdie = x + x - iguess
if bossdie >= max then
bossdie = (max - 1)
elseif bossdie <=min then
bossdie = (min + 1)
end if
end function


msgbox"欢迎挑战终极人工智能（死亡模式1vs1）!! ",,"欢迎使用本程序（vbs脚本= =）! "
randomize(minute(time))
x = int(rnd * 99 + 1)
min = 0
max = 100
aa


sub aa
    do
iguess = inputbox("请输入数字（范围" & min & "~" & max & "）" ,"请猜数")
iguess = iguess + 0
    do while iguess >= max or iguess <=min 
msgbox"你猜的数在范围外!" & vbcelf & "请重新输入。",48,"(error0x000015" & x & ")"
iguess = inputbox("请输入数字（范围" & min & "~" & max & "）" ,"请猜数")
iguess = iguess + 0
    loop

if iguess = x then
msgbox"你猜中了!" & vbcrlf & "点击确定退出。",48,"你输了!"
exit sub
elseif iguess < x then
min = iguess
msgbox"太小了!"
elseif iguess > x then
max = iguess
msgbox"太大了!"
end if

aiguess=bossdie(x,min,max,iguess)
msgbox"终极人工智能猜的是" & aiguess
if aiguess = x then
msgbox"终极人工智能猜中了!" & vbcrlf & "点击确定退出。",48,"你赢了!"
exit sub
elseif aiguess < x then
min = aiguess
msgbox"太小了!"
elseif aiguess > x then
max = aiguess
msgbox"太大了!"
end if
    loop

end sub