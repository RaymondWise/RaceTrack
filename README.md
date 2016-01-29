# RaceTrack
[![Code Review](http://www.zomis.net/codereview/shield/?qid=100714)](http://codereview.stackexchange.com/q/100714/75587) 

A game created in response to [a challenge](http://meta.codereview.stackexchange.com/questions/5623/august-2015-community-challenge?cb=1) on http://codereview.stackexchange.com/

Review of this game:
http://codereview.stackexchange.com/q/100714/75587

Description of game:

In the game of Racetrack, cars race around a track bounded by two concentric closed loops drawn on a square grid 

Each player has a car at an integer position (x,y) on the grid with a velocity vector (vx,vy) that starts at (0,0). Players take turns to move their cars. A move consists of:

 1. updating the velocity vector by adding −1, 0, or +1 to each component;

 2. moving the car to (x+vx,y+vy).
