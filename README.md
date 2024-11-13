# Breakout Clone
This is a proof of concept for a Breakout/Arkanoid like clone in MS Access using VBA.  As of November 2024, this is a draft and while it works, there are some bugs that need to be addressed. I also made a YouTube video that demonstrates its use: https://www.youtube.com/watch?v=GLEZmrcn-Vc

# Description
This project is a clone of the classic game Breakout built using Microsoft Access forms and VBA. It features a paddle, a bouncing ball, and destructible blocks. The player controls the paddle with the keyboard to bounce the ball and destroy all blocks to win the game.

# Features
- Paddle Movement: Controlled using the left and right arrow keys.
- Ball Movement: The ball moves in a two-dimensional space, bouncing off walls, the paddle, and blocks.
- Collision Detection: Detects collisions between the ball and walls, paddle, and blocks.
- Scorekeeping: Tracks the player's score as blocks are destroyed.
- High Score Storage: Saves the highest score using a database table.
- Game Over and Winning Conditions: Displays messages when the player wins or loses.
- Straight-Up Detection: Detects if the ball is moving straight up for too long to avoid getting stuck (part of one of the bugs)

# Variables
- paddleSpeed: Controls the speed of the paddle movement.
- ballSpeedX and ballSpeedY: Control the speed and direction of the ball's movement.
- ballLeft and ballTop: Store the ball's current position.
- gameStarted: Boolean indicating if the game has started.
- score: Tracks the player's score.
- initialBallLeft and initialBallTop: Store the initial position of the ball.
- straightUpCounter and straightUpThreshold: Used to detect if the ball is stuck moving straight up.

# Main Subroutines
- Form_Load: Initializes variables and sets up the game state.
- start_Click: Starts the game and resets the score, ball position, and blocks.
- Form_KeyDown: Handles paddle movement based on keyboard input.
- Form_Timer: Updates the game state on each timer tick, calling subroutines for ball movement, collision detection, and win/loss conditions.
- MoveBall: Updates the ball's position.
- CheckCollisions: Detects collisions between the ball and walls, paddle, and blocks.
- ResetBlocks: Makes all blocks visible (resets game state).
- CheckHighScore: Updates the high score if the player's score exceeds the existing high score.
- CheckStraightUp: Stops the game if the ball is stuck moving straight up.
- CheckWin: Checks if all blocks are destroyed.

# Code Explanation

## 1. Global Variables

```
Dim paddleSpeed As Integer
Dim ballSpeedX As Single
Dim ballSpeedY As Single
Dim ballLeft As Single
Dim ballTop As Single
Dim gameStarted As Boolean
Dim score As Integer
Dim initialBallLeft As Single
Dim initialBallTop As Single
Dim straightUpCounter As Integer
Dim straightUpThreshold As Integer
```
See above for definitions

## 2. Form_Load Event

```
Private Sub Form_Load()
    paddleSpeed = 15 ' Initial paddle speed
    ballSpeedX = 6 ' Initial horizontal ball speed
    ballSpeedY = 6 ' Initial vertical ball speed
    Me.TimerInterval = 0 ' Timer is initially stopped
    ballLeft = Me.ball.Left
    ballTop = Me.ball.Top
    initialBallLeft = ballLeft ' Store initial ball position
    initialBallTop = ballTop ' Store initial ball position
    gameStarted = False ' Game state is initially stopped
    score = 0
    Me.scoreboard.Value = score ' Initialize scoreboard to 0
    straightUpCounter = 0
    straightUpThreshold = 200 ' 200 timer ticks before triggering straight-up detection
End Sub
```
- Initializes variables for paddle and ball speed, ball position, and the game state.
- The TimerInterval is set to 0 to ensure the game does not run automatically on form load.
- The initial positions of the ball are saved to enable resetting later.
- straightUpCounter and straightUpThreshold are used to handle situations where the ball might get "stuck" moving vertically.

## 3. Starting the game with start_click

```
Private Sub start_Click()
    If Not gameStarted Then
        Me.TimerInterval = 10 ' Start the game by setting timer interval (10 ms)
        gameStarted = True ' Game is now running
        score = 0 ' Reset the score
        Me.scoreboard.Value = score ' Display initial score
        ballLeft = initialBallLeft ' Reset ball position
        ballTop = initialBallTop ' Reset ball position
        Me.ball.Move ballLeft, ballTop ' Move the ball to its initial position
        ballSpeedX = 6 ' Reset ball horizontal speed
        ballSpeedY = 6 ' Reset ball vertical speed
        ResetBlocks ' Reset visibility of all blocks
        straightUpCounter = 0 ' Reset straight-up counter
    End If
End Sub
```

- Starts the game by setting TimerInterval to 10 milliseconds.
- Resets the score, ball position, speed, and visibility of blocks.
- straightUpCounter is reset to ensure smooth gameplay.

## 4. Handling Paddle Movement

```
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If gameStarted Then
        Select Case KeyCode
            Case vbKeyLeft
                If Me.cursor.Left > Me.boxborder.Left Then
                    Me.cursor.Left = Me.cursor.Left - paddleSpeed * 8 ' Move paddle to the left
                    If Me.cursor.Left < Me.boxborder.Left Then
                        Me.cursor.Left = Me.boxborder.Left ' Keep paddle within bounds
                    End If
                End If
            Case vbKeyRight
                If Me.cursor.Left < (Me.boxborder.Left + Me.boxborder.Width - Me.cursor.Width) Then
                    Me.cursor.Left = Me.cursor.Left + paddleSpeed * 8 ' Move paddle to the right
                    If Me.cursor.Left > (Me.boxborder.Left + Me.boxborder.Width - Me.cursor.Width) Then
                        Me.cursor.Left = Me.boxborder.Left + Me.boxborder.Width - Me.cursor.Width ' Keep paddle within bounds
                    End If
                End If
        End Select
    End If
End Sub
```
- Moves the paddle left or right based on key presses (vbKeyLeft or vbKeyRight).
- Prevents the paddle from moving out of bounds by enforcing limits.

## 5. Timer Event

```
Private Sub Form_Timer()
    On Error Resume Next ' Resume execution on error
    MoveBall
    If Err.Number <> 0 Then
        MsgBox "Error occurred: " & Err.Description, vbExclamation, "Error"
        Err.Clear ' Clear the error
    End If
    CheckCollisions
    If Err.Number <> 0 Then
        MsgBox "Error occurred: " & Err.Description, vbExclamation, "Error"
        Err.Clear ' Clear the error
    End If
    CheckStraightUp
    If Err.Number <> 0 Then
        MsgBox "Error occurred: " & Err.Description, vbExclamation, "Error"
        Err.Clear ' Clear the error
    End If
    CheckWin ' Check if all blocks are destroyed
    If Err.Number <> 0 Then
        MsgBox "Error occurred: " & Err.Description, vbExclamation, "Error"
        Err.Clear ' Clear the error
    End If
    On Error GoTo 0 ' Turn off error handling
End Sub
```

- Called at intervals specified by TimerInterval.
- Calls subroutines to handle ball movement, collision detection, and win conditions.
- Uses On Error Resume Next to prevent crashes from unexpected errors.

## 6. Moving the Ball

```
Private Sub MoveBall()
    ballLeft = ballLeft + ballSpeedX * 3 ' Update horizontal position
    ballTop = ballTop + ballSpeedY * 3 ' Update vertical position
    Me.ball.Move ballLeft, ballTop ' Update ball position on form
End Sub
```

- Updates the position of the ball based on its speed.
- Multiplying ballSpeedX and ballSpeedY by 3 increases movement speed.

## 7. Collision Detection

```
Private Sub CheckCollisions()
    ' Wall collision
    If ballLeft <= Me.boxborder.Left Or ballLeft >= (Me.boxborder.Left + Me.boxborder.Width - Me.ball.Width) Then
        ballSpeedX = -ballSpeedX ' Reverse horizontal direction
    End If
    If ballTop <= Me.boxborder.Top Then
        ballSpeedY = -ballSpeedY ' Reverse vertical direction
    End If

    ' Paddle collision
    If (ballTop + Me.ball.Height >= Me.cursor.Top) And _
       (ballLeft + Me.ball.Width >= Me.cursor.Left) And _
       (ballLeft <= Me.cursor.Left + Me.cursor.Width) Then
        ballSpeedY = -ballSpeedY ' Reverse vertical direction
        ' Adjust ball speed based on paddle hit position
        Dim hitPosition As Double
        hitPosition = (ballLeft + Me.ball.Width / 2) - (Me.cursor.Left + Me.cursor.Width / 2)
        Dim normalizedHitPosition As Double
        normalizedHitPosition = hitPosition / (Me.cursor.Width / 2)
        ballSpeedX = normalizedHitPosition * 10 ' Adjust horizontal speed
        If Abs(ballSpeedX) < 2 Then
            ballSpeedX = Sgn(ballSpeedX) * 2 ' Prevent ball from moving too slowly
        End If
    End If

    ' Block collision detection and score update
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If ctrl.Tag = "block" And ctrl.Visible Then
            If (ballTop <= (ctrl.Top + ctrl.Height)) And _
               (ballLeft + Me.ball.Width >= ctrl.Left) And _
               (ballLeft <= (ctrl.Left + ctrl.Width)) Then
                ballSpeedY = -ballSpeedY
                ctrl.Visible = False ' Hide the block
                ballSpeedX = ballSpeedX * 1.1 ' Slight speed increase
                ballSpeedY = ballSpeedY * 1.1 ' Slight speed increase
                score = score + 10 ' Increment score
                Me.scoreboard.Value = score
                straightUpCounter = 0 ' Reset straight-up counter
                Exit For ' Exit loop once a collision is detected
            End If
        End If
    Next ctrl

    ' Check if ball falls below paddle (game over)
    If ballTop > (Me.boxborder.Top + Me.boxborder.Height) Then
        CheckHighScore
        Me.TimerInterval = 0 ' Stop game
        gameStarted = False
        MsgBox "Game Over! Your score is " & score, vbExclamation, "Game Over"
    End If
End Sub
```

- Handles collisions with walls, the paddle, and blocks.
- Adjusts ball speed and direction based on collision position.
- Updates score and visibility of blocks.

## 8. Resetting the Blocks (i.e., making them visible)
```
Private Sub ResetBlocks()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If ctrl.Tag = "block" Then
            ctrl.Visible = True ' Make all blocks visible
        End If
    Next ctrl
End Sub
```

- Resets all blocks by making them visible

## 9. Checking for the high score in the "Scores" table

```
Private Sub CheckHighScore()
    Dim rs As DAO.Recordset
    Dim highScore As Integer
    Dim playerName As String

    ' Open scores table
    Set rs = CurrentDb.OpenRecordset("scores")

    ' Find the high score
    If Not rs.EOF Then
        rs.MoveFirst
        highScore = rs.Fields("score").Value
        Do Until rs.EOF
            If rs.Fields("score").Value > highScore Then
                highScore = rs.Fields("score").Value
            End If
            rs.MoveNext
        Loop
    Else
        highScore = 0
    End If

    ' Save new high score
    If score > highScore Then
        playerName = InputBox("Congratulations, new high score! Enter your name:", "New High Score")
        If playerName <> "" Then
            rs.AddNew
            rs.Fields("player").Value = playerName
            rs.Fields("score").Value = score
            rs.Update
        End If
    End If

    ' Close recordset
    rs.Close
    Set rs = Nothing
End Sub
```

- Checks if the current score is a new high score and appends the values to the Scores table if it is.

## 10. Checking for Ball Stuck Movement

```
Private Sub CheckStraightUp()
    If Abs(ballSpeedX) < 2 And ballSpeedY < 0 Then
        straightUpCounter = straightUpCounter + 1
        If straightUpCounter >= straightUpThreshold Then
            Me.TimerInterval = 0
            gameStarted = False
            MsgBox "Game Over! The ball was stuck moving straight up.", vbExclamation, "Game Over"
        End If
    Else
        straightUpCounter = 0
    End If
End Sub
```

- Stops the game if the ball moves straight up for too long.
- This is a temporary fix to one of the bugs I've encountered

## 11. Checking win condition

```
Private Sub CheckWin()
    Dim allBlocksDestroyed As Boolean
    Dim ctrl As Control
    allBlocksDestroyed = True
    
    For Each ctrl In Me.Controls
        If ctrl.Tag = "block" And ctrl.Visible Then
            allBlocksDestroyed = False
            Exit For
        End If
    Next ctrl
    
    If allBlocksDestroyed Then
        Me.TimerInterval = 0
        gameStarted = False
        MsgBox "Congratulations! All blocks destroyed! Your score is " & score, vbInformation, "You Win!"
    End If
End Sub
```

- Stops the game and shows a winning message when all blocks are destroyed.

# Current Issues
- Ball flickering, much worse with images
- Ball will get stuck when it hits the paddle at an almost completion horizontal angle (hard to trigger)
- Ball would get stuck and just go vertically up and down (hence the aforementioned temporary fix)

# Usage
This is a proof of concept and you are free to use as you wish. Please feel free to tweak it and make it your own.  I would be grateful if you could credit Too Long; Didn't Watch Tutorials or TLDW_Tutorials for any future versions of this. I don't know that I will completely finish this project, but it functions and is a good project for someone else to take over. 
