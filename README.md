# VBA
엑셀에서 매크로로 사용할 수 있는 VBA의 기본 문법과
유용한 구문, 응용 구문 몇 가지를 정리해두었다.

* 한셀에서 공부한 내용 옮기는 거라 엑셀이란 다른 부분 찾으면 수정 필요
* 강의서가 아니라 간단한 기록이기 때문에 프로그래밍 기본기 없이 보기는 어려움

## 목차
[1. 기본 정리](#기본-정리)  
[2. 유용한 구문](#유용한-구문)    
[3. 응용](#응용)    


# 기본 정리
## 1. 프로시저

### Sub
```
Sub 서브루틴명 [(인자1, 인자2, ...)]
    명령문
End Sub
```
#### 종료
```
Exit Sub
```
#### 호출
```
Call 서브루틴명
서브루틴명 [(인자1, 인자2, ...)]
```
### Function
```
Function 함수명 [(인자1, 인자2, ...)]
    명령문
    [함수명 = 식]
End Function
```
#### 종료
```
Exit Function
```
#### 호출
```
Call 함수명
함수명 [(인자1, 인자2, ...)]
```
인자가 없으면 Call을 붙이고 없으면 안 붙여야 됐던 걸로 기억함.
확인하면 수정 필요

## 2. 주석
```
' 주석내용
```

## 3. 데이터 형식
#### 타입 확인
```
Typename(변수)
```
타입 알려줌
Null, Boolean, Byte, Integer .... 종류는 특별할 것 없음

## 4. 연산자
### 산술, 연결, 비교 연산자
특별할 것 없이 +-<> 이용하면 됨
### 논리 연산자
Not, And, Or, Eqv, Xor 사용

## 5. 분기문, 반복문 등
### If
```
If 조건1 Then
    명령문1
[ElseIf 조건2 Then
    명령문2
Else
    명령문3)
End If
```
3항 연산자가 되던가? 확인하면 추가

### Case
```
Select Case 표현식
    Case 값1 [,값3, 값4..] / /값 여러개 표현'/
        명령문1
    [Case 값2 
        명령문2
    Case Else
        명령문3 ]
End Select
```

### For
```
For 변수 = 값1 To 값2 [step 값3]
    명령문
Next
```
```
For Each 요소 In 배열
    명령문
Next
```
For문 내에서 요소 값 사용가능

#### 종료
```
Exit For
```

### Do While
```
Do While 조건
    명령문
Loop
```
조건이 참인 동안 명령문 반복
```
Do
    명령문
Loop While 조건
```
루프를 일단 한번은 실행
#### 종료
```
Exit Do
```

### Do Until
```
Do Until 조건
    명령문
Loop
```
조건이 참인 동안 명령문 반복
```
Do
    명령문
Loop Until 조건
```
루프를 일단 한번은 실행

#### 종료
```
Exit Do
```

### While
```
While 조건
    명령문
Wend
```
이게 되나? 안 써봐서 모르겠는데 나중에 안되면 빼기

### With
```
With 개체
    명령문
    .속성
    .메소드
End With
```
이중으로 사용 가능하다.

## 6. 에러 제어문
```
On Error Resume Next '무시하고 넘어가기'
 ```
 ```
On Error GoTo 0 '기본값' / /에러를 무시하고 다음 줄로 넘어가게 하는 코드/
```
잘만 활용하면 코드 쓰기가 편해진다.

예시)
주변 셀을 확인하는 함수를 반복 시킬 때 (특히 재귀 함수에서)
행, 열이 1보다 작은 값을 참조하면 에러가 뜨는데
매번 범위를 확인하는 것보다 에러일 때 무시하고 끝내버리는 게
코드도 간단하고 연산도 줄어듦

# 유용한 구문
## 1. 속도 향상
### 매크로가 시작될 때
```
Application.ScreenUpdating = False '화면 업데이트
Application.Calculation = xlCalculationManual '수식 자동계산
```
### 매크로가 끝날 때
```
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
```
보통 몇배 이상 차이가 난다.
그냥 필수로 넣자.

# 응용
## SelectionChange로 버튼 만들기
### 1. 버튼 만들기
매크로를 만들 때 단추나 도형을 넣어서 매크로를 연결해주는 게 보통이지만
SelectionChange를 이용하면 시트의 셀을 눌렀을 때 매크로가 실행되게 할 수 있다.
또 여러 셀은 선택한 경우에도 대응할 수 있으니 여러모로 응용하기 좋다.

### 2. SelectionChange 준비하기
1) 매크로 편집창에서 모듈을 선택하지 말고 시트(Sheet1, ...)나 워크북(현재_통합_문서)를 선택
2-1) 시트는 상단의 (일반) - (선언)을 클릭해서 Worksheet - SelectionChange 를 클릭
2-2) 워크북이면 Workbook - SheetSelectionChange 를 클릭

#### 시트에서
```
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
```
-----------------------------------------------------------------------
#### 워크북에서
```
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)

End Sub
```
3) 위처럼 Sub가 자동으로 생기면 준비 완료

* SelectionChange 외에도 유용하게 사용 가능한 함수가 많으니 자기가 필요한 게 있는지 잘 찾아보자

### 3.  SelectionChange로 버튼 만들기
선택된 영역을 Target 이란이름의 Range형으로 참조할 수 있다.(Selection으로도 가능)

#### 기본1 - 셀 위치를 이용
```
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Address = Cells(2, 11).Address Then
        명령문
    End If
End Sub
```
(2,11)셀을 선택했을 때 명령문이 실행된다.

```
cells(1,1).select
```
명령문 다음에 버튼과 상관 없는 셀로 선택 위치를 옮겨주면
연속으로 같은 버튼을 여러분 누를 수 있다. 거의 필수로 쓰임
병합된 셀이 있을때 첫번째 셀 주소를 참조하는지 병합된 셀 전체를 참조하는지 까먹었는데 확인하면 추가

#### 기본2 - 값을 이용
```
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Address = "버튼명"
        명령문
    End If
End Sub
```
"버튼명"이라고 적힌 셀을 선택했을 때 명령문이 실행된다.
시트에서 버튼 위치를 바꾸고 싶을때 코드를 바꾸지 않아도 되서 유용하다.

하지만 버튼명과 같은 내용이 적힌 다른 셀이 있으면 안되니 주의해야한다.
버튼이 있을 수 있는 영역을 지정하는 것으로 해결 가능
버튼에 별다른 제약 사항이 없다면 위 두 가지 방법 모두 Select Case를 사용하면 더 간단해진다.

* 추가 예정
- 첫번째 셀의 행, 열 값과, 선택된 영역의 행, 열 크기 이용하기 + 네 부분 모서리 구하기 가능
- 버튼, 선택한 영역 위치나 크기 제한하기
- 좀 복잡하게 응용하면 도트 찍는 그림판 만들기 가능
