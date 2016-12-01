# The 2016-17 NBA Season Ranking System</br>
##Goal

When it comes to betting sports, many people believe, you have to predict the final spread of the game, if you want to win. This would make winning simplier, but since predicting the score of a game is nearly impossible, due to factors that take place during game, we need a different goal. The normal goal for betters is to win money, and the only way to do that is pick the winning side of the Vegas line. So, if we understand Vegas' system for producing a line, we could understand how to be on the winning side of the line.

####Vegas Goal
Many people think that Vegas is trying to predict the final score of the game, but this isn't true. Vegas has a algorithm that outputs a spread value for each game, then open bets to the public. As people place bets, Vegas will shift the line when neccessary so that it can keep money even on both sides of the line. When the game is over, Vegas will take the money from the losing side, and use it to pay the winners, while Vegas collects the juice. Knowing this, it would be fair to say, that in Vegas' algorithm, although it does try to forecast a reasonable spread, they also consider the spread value that people would expect the game to be. Therefore, betting habits would also be a part of the Vegas algorithm.

####What Is A Win Algorithm?
If our goal is to win money betting, then having a minimum win percentage wouldn't classify an algorithm as successful. Everytime you bet the spread, you have the added risk of the juice, -110. Before you even place your bet, Vegas has an expected win percentage os 52.4%. There is a simple formula you can use to breakdown the moneyline and determine what is Vegas' expected win percentage.

<strong>Negative Money Line = [-(Money Line)] / [- (Money Line) + 100]</strong></br>

So, to determine Vegas' expected win percent needed for the spread we can just input the -110 into the formula, and we'll output the expected win percent.</br>

<strong>Vegas' Expected Win = [-(-110)]/[-(-110)+100] = 0.5238 or 52.3%</strong></br>

If you only play the Vegas spread, then you would have to win a minimum of 52.5%, to gain a profit. Every bet can have different moneylines, representing the different expected wins for Vegas.</br></br>

<strong>Positive Money Line = 100 / (Money Line + 100)</strong></br>

##Strategy

First, You want to create a data table with every game of the season. This table will hold the data for scores, lines, expected values, and outcomes represented as 0 for a loss and 1 for a win/push. Then I can easily take the sum of the column divided by the count of the column to return a win percentage.</br>

Second, you want to create an equation that produces an expected spread value that is consistentent and reasonable. The purpose of this spread is to represent the knowledge of the average better. It is used to determine what side of the line the average better will bet on.</br>

Third,You want to create a new worksheet that will hold multiple tables that will show the probability of different events being correct.

##Equation

<!-- Offensive Rating -->

<img src="http://www.sciweavers.org/tex2img.php?eq=Offensive%20Rating%20%3D%20%28Avg.Pts.%2FPossession%29%2A100&bc=White&fc=Black&im=jpg&fs=12&ff=arev&edit=0" align="center" border="0" alt="Offensive Rating = (Avg.Pts./Possession)*100" width="412" height="19" />

<!-- Pace -->

<img src="http://www.sciweavers.org/tex2img.php?eq=Pace%20Weight%20%3D%20%28Avg.%20Possessions%2FGame%29%20%2F%20100&bc=White&fc=Black&im=jpg&fs=12&ff=arev&edit=0" align="center" border="0" alt="Pace Weight = (Avg. Possessions/Game) / 100" width="386" height="19" />

<!-- Defensive Rating -->

<img src="http://www.sciweavers.org/tex2img.php?eq=Defensive%20Rating%20%3D%20%28Avg.Pts.%2FPossession%29%2A100&bc=White&fc=Black&im=jpg&fs=12&ff=arev&edit=0" align="center" border="0" alt="Defensive Rating = (Avg.Pts./Possession)*100" width="414" height="19" /></br>

The defensive differential of all teams must average out to 0. I haven't seen anything larger than +/-15, which just mean on average we allow no more than 15 points on average, or prevent 15 points on average.</br>

<img src="http://www.sciweavers.org/tex2img.php?eq=Defensive%20Differential%20%3D%20%28TeamDefensiveRating%29-%28LeagueAvg.DefensiveRating%29&bc=White&fc=Black&im=jpg&fs=12&ff=arev&edit=0" align="center" border="0" alt="Defensive Differential = (TeamDefensiveRating)-(LeagueAvg.DefensiveRating)" width="712" height="19" />

<!-- Home Edge -->

There is a belief the the home edge of all teams is 3 points. This isn't entirely true. The average of all teams home edge will always equal between 2.5 and 3.5.</br>

<img src="http://www.sciweavers.org/tex2img.php?eq=Home%20Edge%20%3D%20%28Team%20Avg.%20Home%20Pts.%29%20-%20%28Team%20Avg.%20Away%20Pts.%29&bc=White&fc=Black&im=jpg&fs=12&ff=arev&edit=0" align="center" border="0" alt="Home Edge = (Team Avg. Home Pts.) - (Team Avg. Away Pts.)" width="508" height="19" />

<!-- Strength of Schedule -->

<!-- Team Rating -->

##Observations
