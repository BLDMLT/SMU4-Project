/**
* Returns the first name as preferred name
* @param x Full name "First_name Middle_name Last_name"
* @customfunction
*/
function firstPreferredName(x) {
  var arr = x.split(" ");
  return arr[0];
}

/**
* Returns the last name as preferred name
* @param x Full name "First_name Middle_name Last_name"
* @customfunction
*/
function lastPreferredName(x) {
  var arr = x.split(" ");
  return arr[arr.length - 1];
}

/**
* Returns the email FullName@maptek.com
* @param x Full name "First_name Middle_name Last_name"
* @customfunction
*/
function emailFullName(x) {
  var temp = x.replace(/ /g,"");
  var email = temp + "@maptek.com";
  email = email.toLowerCase();
  return email;
}


/**
* Returns the Random Food
* @customfunction
*/
function randomFood() {
  var food = ["Pizza","Burger","Hot Dog","Pasta","Grilled Chicken","Raps"];
  var temp = Math.floor(Math.random() * food.length)
  return food[temp];
}

/**
* Returns the Random Rating
* @customfunction
*/
function randomRating() {
  var temp = Math.floor(Math.random() * 10) + 1;
  return temp;
}

