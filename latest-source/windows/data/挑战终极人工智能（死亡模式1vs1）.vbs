
function bossdie(x , min , max, iguess  )
bossdie = x + x - iguess
if bossdie >= max then
bossdie = (max - 1)
elseif bossdie <=min then
bossdie = (min + 1)
end if
end function


msgbox"��ӭ��ս�ռ��˹����ܣ�����ģʽ1vs1��!! ",,"��ӭʹ�ñ�����vbs�ű�= =��! "
randomize(minute(time))
x = int(rnd * 99 + 1)
min = 0
max = 100
aa


sub aa
    do
iguess = inputbox("���������֣���Χ" & min & "~" & max & "��" ,"�����")
iguess = iguess + 0
    do while iguess >= max or iguess <=min 
msgbox"��µ����ڷ�Χ��!" & vbcelf & "���������롣",48,"(error0x000015" & x & ")"
iguess = inputbox("���������֣���Χ" & min & "~" & max & "��" ,"�����")
iguess = iguess + 0
    loop

if iguess = x then
msgbox"�������!" & vbcrlf & "���ȷ���˳���",48,"������!"
exit sub
elseif iguess < x then
min = iguess
msgbox"̫С��!"
elseif iguess > x then
max = iguess
msgbox"̫����!"
end if

aiguess=bossdie(x,min,max,iguess)
msgbox"�ռ��˹����ܲµ���" & aiguess
if aiguess = x then
msgbox"�ռ��˹����ܲ�����!" & vbcrlf & "���ȷ���˳���",48,"��Ӯ��!"
exit sub
elseif aiguess < x then
min = aiguess
msgbox"̫С��!"
elseif aiguess > x then
max = aiguess
msgbox"̫����!"
end if
    loop

end sub