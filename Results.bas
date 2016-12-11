Attribute VB_Name = "Results"
'This function is used to count the number of wins in the win column based on division passed
'If the divisional_col equals TRUE, and wins_col equals 1, and division_col equals division
'If these conditions return TRUE then add 1 to a counter to be returned as the number of wins for the passed division
Function DivisionWins(division As String, wins_col As Range, divisional_col As Range, division_col As Range)
    
    'Counter for each loop
    Dim current_element As Integer: current_element = 1
    Dim total_wins As Integer: total_wins = 0
    
    'For each element in division_col if value = TRUE, then finish loop, if FALSE continue
    For Each div In divisional_col
        
        'If current game is a divisional
        If div.Value = True Then
            
            'If the divisional game is in the division being looked for
            If division_col(current_element).Value = division Then
                
                If wins_col(current_element) <> "" Then
                'Then increment the win counter
                total_wins = total_wins + wins_col(current_element).Value
                
                End If
            End If
        End If
        
        'Increment counter
        current_element = current_element + 1
    Next
    
    DivisionWins = total_wins
End Function

'This function is used to count the number of games in the win column based on division passed
Function DivisionGames(division As String, wins_col As Range, divisional_col As Range, division_col As Range)
    
    'Counter for each loop
    Dim current_element As Integer: current_element = 1
    Dim total_games As Integer: total_games = 0
    
    'For each element in division_col if value = TRUE, then finish loop, if FALSE continue
    For Each div In divisional_col
    
        'If current game is a divisional
        If div.Value = True Then
        
            'If the divisional game is in the division being looked for
            If division_col(current_element).Value = division Then
                
                'If the game has been recorded
                If wins_col(current_element) <> "" Then
                    'Then increment the win counter
                    total_games = total_games + 1
                
                End If
            End If
        End If
        
        'Increment counter
        current_element = current_element + 1
    Next
    
    DivisionGames = total_games
End Function

Function TeamGames(team As String, home_col As Range, away_col As Range, wins_col As Range, team_col As Range, side As Integer)
    'Iteration for for loop
    Dim cur_elm As Integer: cur_elm = 1
    Dim num_games As Integer: num_games = 0
    Dim home_value As String: home_value = "HOME"
    Dim away_value As String: away_value = "AWAY"
    
    If side = 1 Then
        home_value = "AWAY"
        away_value = "HOME"
    End If
    
    'For each element in home column check if the team name is the team name passed
    For Each home In home_col
        'If the current home team is the team looking for
        If home.Value = team Then
            'If there is a value in the team column
            If team_col(cur_elm) = home_value Then
                'If there is a value in the wins column
                If wins_col(cur_elm) <> "" Then
                    'Increment the number of games
                    num_games = num_games + 1
                End If
            End If
        End If
        'Increment current element
        cur_elm = cur_elm + 1
    Next
    
    cur_elm = 1
    'For each element in away column check if the team name is the team name passed
    For Each away In away_col
        'If the current away team is the team looking for
        If away.Value = team Then
            'If there is a value in the team column
            If team_col(cur_elm) = away_value Then
                'If there is a value in the wins column
                If wins_col(cur_elm) <> "" Then
                    'Increment the number of games
                    num_games = num_games + 1
                End If
            End If
        End If
        'Increment current element
        cur_elm = cur_elm + 1
    Next
    
    TeamGames = num_games
    
End Function

Function TeamWins(team As String, home_col As Range, away_col As Range, wins_col As Range, team_col As Range, side As Integer)
    'Iteration for for loop
    Dim cur_elm As Integer: cur_elm = 1
    Dim num_wins As Integer: num_wins = 0
    Dim home_value As String: home_value = "HOME"
    Dim away_value As String: away_value = "AWAY"
    
    If side = 1 Then
        home_value = "AWAY"
        away_value = "HOME"
    End If
    
    'For each element in home column check if the team name is the team name passed
    For Each home In home_col
        'If the current home team is the team looking for
        If home.Value = team Then
            'If there is a value in the team column
            If team_col(cur_elm) = home_value Then
                'If there is a value in the wins column
                If wins_col(cur_elm) <> "" Then
                    If side = 1 Then
                        'Increment the number of games
                        num_wins = num_wins + (1 - wins_col(cur_elm).Value)
                    Else
                        'Increment the number of games
                        num_wins = num_wins + wins_col(cur_elm).Value
                    End If
                End If
            End If
        End If
        'Increment current element
        cur_elm = cur_elm + 1
    Next
    
    cur_elm = 1
    'For each element in away column check if the team name is the team name passed
    For Each away In away_col
        'If the current away team is the team looking for
        If away.Value = team Then
            'If there is a value in the team column
            If team_col(cur_elm) = away_value Then
                'If there is a value in the wins column
                If wins_col(cur_elm) <> "" Then
                    If side = 1 Then
                        'Increment the number of games
                        num_wins = num_wins + (1 - wins_col(cur_elm).Value)
                    Else
                        'Increment the number of games
                        num_wins = num_wins + wins_col(cur_elm).Value
                    End If
                End If
            End If
        End If
        'Increment current element
        cur_elm = cur_elm + 1
    Next
    
    TeamWins = num_wins
    
End Function

Function HomeWins(wins_col As Range, team_col As Range)
    'Iteration for for loop
    Dim cur_elm As Integer: cur_elm = 1
    Dim num_wins As Integer: num_wins = 0
    
    'For each element in home column check if the team name is the team name passed
    For Each home In team_col
            'If there is a value in the team column
            If home = "HOME" Then
                'If there is a value in the wins column
                If wins_col(cur_elm) <> "" Then
                    'Increment the number of games
                    num_wins = num_wins + wins_col(cur_elm).Value
                End If
            End If
            
        'Increment current element
        cur_elm = cur_elm + 1
    Next
    
    HomeWins = num_wins
    
End Function

Function AwayWins(wins_col As Range, team_col As Range)
    'Iteration for for loop
    Dim cur_elm As Integer: cur_elm = 1
    Dim num_wins As Integer: num_wins = 0
    
    'For each element in home column check if the team name is the team name passed
    For Each away In team_col
            'If there is a value in the team column
            If away = "AWAY" Then
                'If there is a value in the wins column
                If wins_col(cur_elm) <> "" Then
                    'Increment the number of games
                    num_wins = num_wins + wins_col(cur_elm).Value
                End If
            End If
            
        'Increment current element
        cur_elm = cur_elm + 1
    Next
    
    AwayWins = num_wins
    
End Function

Function HomeGames(wins_col As Range, team_col As Range)
    'Iteration for for loop
    Dim cur_elm As Integer: cur_elm = 1
    Dim num_games As Integer: num_games = 0
    
    'For each element in home column check if the team name is the team name passed
    For Each home In team_col
            'If there is a value in the team column
            If home = "HOME" Then
                'If there is a value in the wins column
                If wins_col(cur_elm) <> "" Then
                    'Increment the number of games
                    num_games = num_games + 1
                End If
            End If
            
        'Increment current element
        cur_elm = cur_elm + 1
    Next
    
    HomeGames = num_games
    
End Function

Function AwayGames(wins_col As Range, team_col As Range)
    'Iteration for for loop
    Dim cur_elm As Integer: cur_elm = 1
    Dim num_games As Integer: num_games = 0
    
    'For each element in home column check if the team name is the team name passed
    For Each away In team_col
            'If there is a value in the team column
            If away = "AWAY" Then
                'If there is a value in the wins column
                If wins_col(cur_elm) <> "" Then
                    'Increment the number of games
                    num_games = num_games + 1
                End If
            End If
            
        'Increment current element
        cur_elm = cur_elm + 1
    Next
    
    AwayGames = num_games
    
End Function
