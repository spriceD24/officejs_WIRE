/* eslint-disable @typescript-eslint/no-unused-vars */
/* global console setInterval, clearInterval */

/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

/**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param {string} message String to write.
 * @returns String to write.
 */
function logMessage(message) {
  console.log(message);

  return message;
}

/**
 * This custom function sends a request to the server with text and number parameters, and an optional delay.
 * @customfunction
 * @param {string} text The input text.
 * @param {number} number The multiplier number.
 * @param {number} [delay] Optional delay in seconds before returning the result.
 * @returns {Promise<number|string>} The result from the API or an error message.
 */
async function UBSFIN(text, number, delay) {
  try {
    console.log("UBSFIN");

    // Construct the API URL with query parameters
    let url = `http://localhost:5000/wire?text=${encodeURIComponent(text)}&number=${number}`;

    // Add the delay parameter to the URL if provided and is a positive number
    if (typeof delay !== 'undefined' && delay > 0) {
      url += `&delay=${delay}`;
    }

    console.log(`Calling URL: ${url}`);

    // Make the API request
    const response = await fetch(url);

    // Check if the response is ok (status 200)
    if (!response.ok) {
      throw new Error(`Error: ${response.statusText}`);
    }

    // Parse the JSON response
    const data = await response.json();

    // Return the result from the API
    return data.result;
  } catch (error) {
    // Handle any errors
    console.error(error);
    return "Error calling API";
  }
}

/**
 * This custom function returns "Hello, World!".
 * @customfunction
 * @returns A string that says "Hello, World!".
 */
function HELLO() {
  return "Hello, World!";
}


// You must include this line to make the function available in Excel.
CustomFunctions.associate("UBSFIN", UBSFIN);
CustomFunctions.associate("HELLO", HELLO);