# 셀값이 바뀌면 페이지를 구분

## 출처

[http://blog.naver.com/rosa0189/60144773920](http://blog.naver.com/rosa0189/60144773920)

## Algorithm
 
1. 데이터가 입력된 마지막 행 구하기

```VBScript
lastRow = Cells(Rows.Count, 1).End(3).Row
```

**Cells**: Cell property 숫자를 인수로 사용할 수 있음. Cells(행번호, 열번호)
**Rows.Count**: 전체 행 개수 **Cells(Rows.Count, 1)**: A열 끝행을 가리킴. 1대신 A를 써도
됨. **End(3)**: xlUp. 선택된 셀에서부터 위로 이동. 1은 좌, 2는 우, 3은 상, 4는 하로 이동.
**Cells(Rows.Count, 1).End(3)**: A열 끝행에서부터 값이 입력된 셀을 만날 때까지 윗셀로 이동.
**Row**: .Row 앞에 온 개체가 위치한 행

위 코드는 A열의 마지막 셀에서부터 위로 올라가다가 데이터가 입력된 셀을 만나면 그 셀의 행번호를 표시. 즉, 데이터가 입력된
마지막 셀의 행번호 반환

2. 마지막 셀부터 거꾸로 올라오며 위의 셀값과 비교

```VBScript
for r = lastRow to header step -1
	if Cells(r, 1) <> Cells(r-1, 1) Then
		
	End If
next r
```

3. 페이지 구분선 삽입

```VBScript
ActiveSheet.HPageBreaks.Add Before:=Rows(10)
ActiveSheet.HpageBreaks.Add Rows(10)
ActiveSheet.HpageBreaks.Add Range("A10")
```

9행과 10행 사이에 페이지 구분선 삽입


## Source

```VBScript
Option Explicit
Sub insert_pageBreaks()

	Dim lastRow As Long                               '전체 행 숫자 넣을 변수
	Dim r As Long                                     '행을 하나씩 줄여갈 변수

	Application.ScreenUpdating = False                '화면 업데이트 (일시)정지
	lastRow = Cells(Rows.Count, 1).End(3).Row     	  '전체데이터 마지막 행번호

	For r = lastRow To 3 Step -1                      'Step - 1이용 아래에서 위로 행줄여가며 반복
		If Cells(r, 2) <> Cells(r - 1, 2) Then        '아래위 값이 다르면
			ActiveSheet.HPageBreaks.Add Cells(r, 1)   '페이지나누기 삽입
		End If
	Next r 
	With ActiveSheet.PageSetup                        '현재시트 페이지 설정
		.PrintTitleRows = "$1:$1"                     '반복 인쇄될 영역(제목 반복 인쇄)
'			.PrintArea = ActiveSheet.UsedRange.Address    '데이터가 있는 영역만을 인쇄영역을 설정
		.PrintArea = ActiveSheet.Range("A1").currentRange.Address	'A1과 인접한 데이터 영역을 인쇄영역으로 설정
	End With

End Sub
```
