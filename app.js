/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

(function () {
  let x = 0;
  let PA = document.getElementById('prince-ali');
  PA.addEventListener('click', (e) => run(e));
  const lyrics = [
"Make way for Prince Ali",
"Say hey! It's Prince Ali",
"Hey! Clear the way in the old Bazaar",
"Hey you!",
"Let us through!",
"It's a bright new star!",
"Oh Come!",
"Be the first on your block to meet his eye!",
"Make way!",
"Here he comes!",
"Ring bells! Bang the drums!",
"Are you gonna love this guy!",
"Prince Ali! Fabulous he!",
"Ali Ababwa",
"Genuflect, show some respect",
"Down on one knee!",
"Now, try your best to stay calm",
"Brush up your sunday salaam",
"The come and meet his spectacular coterie",
"Prince Ali!",
"Mighty is he!",
"Ali Ababwa",
"Strong as ten regular men, definitely!",
"He faced the galloping hordes",
"A hundred bad guys with swords",
"Who sent those goons to their lords?",
"Why, Prince Ali",
"He's got seventy-five golden camels",
"Purple peacocks",
"He's got fifty-three",
"When it comes to exotic-type mammals",
"Has he got a zoo?",
"I'm telling you, it's a world-class menagerie",
"Prince Ali! Handsome is he, Ali Ababwa",
"That physique! How can I speak",
"Weak at the knee",
"Well, get on out in that square",
"Adjust your veil and prepare",
"To gawk and grovel and stare at Prince Ali!",
"He's got ninety-five white Persian monkeys",
"(He's got the monkeys, let's see the monkeys)",
"And to view them he charges no fee",
"(He's generous, so generous)",
"He's got slaves, he's got servants and flunkies",
"(Proud to work for him)",
"They bow to his whim love serving him",
"They're just lousy with loyalty to Ali! Prince Ali!",
"Prince Ali!",
"Amorous he! Ali Ababwa",
"Heard your princess was a sight lovely to see",
"And that, good people, is why he got dolled up and dropped by",
"With sixty elephants, llamas galore",
"With his bears and lions",
"A brass band and more",
"With his forty fakirs, his cooks, his bakers",
"His birds that warble on key",
"Make way for prince Ali!",
  ];

  const inc = () => {
    x = (x + 1) % lyrics.length
  }

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      $('#prince-ali').click(run);
    });
  };

  function run() {
    // const PA = document.getElementById('prince-ali')
    inc()
    PA.innerText = lyrics[x]
    // return Excel.run(function (context) {
    //   #<{(|*
    //    * Insert your Excel code here
    //    |)}>#
    //   return context.sync();
    // });
    
  }

})();
