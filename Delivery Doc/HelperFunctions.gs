/*===========================================================
  HelperFunctions.gs
  ------------------
  
  Functions that can be utilized in more than one project,
  in an effort to not recreate the wheel.
  
  
  LEGEND:
  
    isLeftContainingRight(left, right)
      => returns true if left text contains right text;
         returns false otherwise
  
  
  USE EXAMPLE:
  
    if( isLeftContainingRight(large,small) ) {}
=============================================================
Ralph McCraken, III */



function
isLeftContainingRight(large,small)
{
  if (large == undefined)
    return false;
  
  return large.indexOf(small) >= 0;
}
