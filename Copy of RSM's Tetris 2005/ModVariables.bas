Attribute VB_Name = "ModVariables"
'   __________                       __   _____
'   \________/ _____  /  |_ ______   |_| /  ___|
'      |  |  _/ ___ \_\____\\_  __ \ _ _ \_____
'      |  |   \  ___/   |  | |  | \/ |  | ____ \
'      \__/    \___    |__|  |__|    |__||_____/
'
' _________                               ____________
' _________ T H E   G A M E   B E G I N S ____________
'
' Tetris - A free 2D real time strategy game engine
' Copyright (C) 2002-2003 Mohit Soam
'
'Module Name  --> ModVariables

Option Explicit

'--------------------------------------------------------------------------
'Global variables declarations
'--------------------------------------------------------------------------
Global ScoreIndex As Integer

'--------------------------------------------------------------------------
'Boolean variables declarations
'--------------------------------------------------------------------------
Public GAME_OVER As Boolean 'To Check game Status
Public GamePause As Boolean 'if true the Pause Game
Public START As Boolean 'if true then Draw new Block
Public AVILABLE As Boolean
'if true then next postion is avilable
Public MaxBrick As Boolean
'if true the Next Position is illegel move
'Public CheckStat As Boolean
'Public DownKeyPressed As Boolean
Public LeftBrick As Boolean
'if True then Brick is at rightmost Position
Public RightBrick As Boolean
'if True then Brick is at rightmost Position

'--------------------------------------------------------------------------
'Integer variables declarations
'--------------------------------------------------------------------------
Public Level As Integer
Public TimeCount As Integer
Public Brick As Integer 'To Select Type of Brick
Public Brick1 As Integer 'To Select Type of Next Brick
Public Color As Integer 'To Select Color of Brick
Public Color1 As Integer 'To Select Color of Next Brick
Public BonusCount As Integer
Public BrickCount As Integer
'To count the number of Bricks in a Row if is equals
'to 10 then erase whole row & increase Score

'--------------------------------------------------------------------------
'Array declarations
'--------------------------------------------------------------------------
'To Store current Position of Brick
Public CurrentPosition(5) As Integer
'--- or Public CurrentPosition(0 To 5) As Integer

'To Store Next Position of Brick
Public NextPosition(5) As Integer
'--- or Public NextPosition(0 To 5) As Integer

'To Store Position Pointer of Brick
Public PositionPointer(5) As Integer
'--- or Public PositionPointer(0 To 5) As Integer


