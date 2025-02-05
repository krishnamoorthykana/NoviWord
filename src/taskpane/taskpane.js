/* eslint-disable prettier/prettier */
/* eslint-disable no-undef */
 let speechFlag = false;
Office.onReady(async function (info) {
  displayStartingMessage("Hi, I am your word assistant bot-NoviWord");
  let directLine1 = await initializeDirectLine();
if (info.host === Office.HostType.Word) {
  //let flag=true;
 
document.getElementById("askButton").onclick = async function () {
  const question = document.getElementById("userInput").value;
  if (question) {
    document.getElementById("headerId").style.display = "none"; 
    displayChatMessage(question, '', "User");
      await getBotResponse(directLine1, question);
   
 
  }
};
 
document.getElementById("userInput").addEventListener("keydown", async function (event) {
  if (event.key === "Enter") {
    // Check if Enter key is pressed
    event.preventDefault(); // Prevents the default behavior (like submitting a form)
    document.getElementById("headerId").style.display = "none";
    const question = document.getElementById("userInput").value;
    if (question) {
     
        displayChatMessage(question, '', "User");
      await getBotResponse(directLine1, question);
     
  }
}});
 
// Handle the Insert button click
document.getElementById("insertButton").onclick = async function () {
  const response = document.getElementById("chatWindow").lastChild
    ? document.getElementById("chatWindow").lastChild.innerHTML
    : "";
  if (response) {
    await insertResponseIntoDocument(response);
  }
};

document.getElementById('startSpeechButton').addEventListener('click', function () {
  // Open a pop-up window to handle the speech
  var transcript = null;
  const popup = window.open('speech.html', 'SpeechRecognition', 'width=400,height=300');
speechFlag = true;
  // Listen for messages from the pop-up window
  window.addEventListener("message", async function (event) {
      if (event.origin !== window.location.origin) return; // Security check

      // Get the recognized text from the pop-up
      transcript = event.data;

      // Insert recognized text into user input 
      console.log(transcript);
      document.getElementById("userInput").value = transcript;
      var question = document.getElementById("userInput").value;
    if (question) {
        console.log("Testing inside speech if question");
        displayChatMessage(question, '', "User");
      await getBotResponse(directLine1, question);
     
  }
      popup.close();
  });
});


}
});


function displayStartingMessage(starter) {
  const chatWindow = document.getElementById("chatWindow");
  chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">${starter}</div>`;      
}
// Display user question and bot response in chat window
function displayChatMessage(question, response, role) {
  const chatWindow = document.getElementById("chatWindow");
 
  // Check if response is valid and if attachments exist
  if (response && response.attachments && response.attachments.length > 0) {
    response.attachments.forEach((attachment) => {
      // Check if attachment content has 'buttons' and 'signin' type
      if (attachment.content && attachment.content.buttons && attachment.content.buttons.length > 0) {
        attachment.content.buttons.forEach((button) => {
          if (button.type === "signin") {
            // Create a sign-in button
            const signinButton = document.createElement("button");
            signinButton.innerText = button.title || "Sign In"; // Default title to "Sign In"
            signinButton.classList.add("ms-Button", "ms-Button--primary");
 
            // Open the sign-in URL when the button is clicked
            signinButton.onclick = () => {
              window.open(button.value, "_blank"); // Open the sign-in URL in a new tab
            };
 
            // Display the bot's message
            chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">${attachment.content.text}</div>`;
            chatWindow.appendChild(signinButton); // Add the button after the message
          }
        });
      }
    });
  } else {
    // Regular message display if no attachments
    if (role === "bot") {
      if(response.speak==="Generate"){
        insertResponseIntoDocument(response.text);
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">SOW content generated in document</div>`;
      }
      else if(response.text){
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">${response.text}</div>`; 
        if(speechFlag){ 
        //   ensureVoicesLoaded(() => {
        //     speakText(response.text);
        // });// Example usage:
        speakText(response.text);
        speechFlag = false;  
        }    
      }
    } else {
      if(question){
     
        chatWindow.innerHTML += `<div class="user-wrapper">You</div><div class="message user">${question}</div>`;      }
     
    }
  }
  scrollToBottom();
  // Clear the input field
  document.getElementById("userInput").value = "";
}
 
// Function to insert the response into the Word document
async function insertResponseIntoDocument(response) {
  console.log("Testing insert to doc*********");
  console.log("response*******", response);
  await Word.run(async (context) => {
    console.log("Inside Testing insert to doc*********");
  console.log("Inside response*******", response);
    const body = context.document.body;
    body.insertHtml(response, Word.InsertLocation.end);
    await context.sync();
  });
}
const initializeDirectLine = async function () {
  try {
    const response = await fetch(
      "https://148a369decc3eeda85b913c1e80b9a.da.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cra27_agent123/directline/token?api-version=2022-03-01-preview"
    );
    const data = await response.json();
   
    const directLine = new window.DirectLine.DirectLine({ token: data.token });
   
 
    if (!directLine || !directLine.activity$) {
      throw new Error("DirectLine instance failed to initialize");
    }
 
    directLine
      .postActivity({
        from: { id: "10", name: "User" },
        type: "message",
        text: "Hi",
      })
      .subscribe(
        (id) => console.log("Message sent with ID:", id),
        (error) => console.error("Error sending message:", error)
      );
 
    directLine.activity$.subscribe((activity) => {
      console.log("Testing activity: ", activity);
      console.log("Role", activity.from.role);
      if (activity.type === "message" && activity.from.id !== "10" && !activity.recipient) {
        console.log("Bot Response: ", activity.text);
        displayChatMessage(false, activity, activity.from.role);
       
      }
    });
    return directLine;
  } catch (error) {
    console.error("Error initializing DirectLine:", error);
  }
};
 
const getBotResponse = async function (directLine, question) {
  directLine
    .postActivity({
      from: { id: "10", name: "User" },
      type: "message",
      text: question,
    })
    .subscribe(             // calls the subscription already created in InitializeDirectline method
      (id) => console.log("Message sent with ID:", id),
      (error) => console.error("Error sending message:", error)
    );
 
}
function scrollToBottom() {
  const chatWindow = document.getElementById("chatWindow");
  setTimeout(() => {
    chatWindow.scrollTop = chatWindow.scrollHeight;
  }, 100); // Timeout ensures scroll happens after the new message is rendered
}

// function speakText(text) {
//   console.log("Testing Text to Speech");
//   let voices = window.speechSynthesis.getVoices();
//   console.log("Voices******", voices);
//   let femaleVoice = voices.find(voice => voice.name.includes("Female") || 
//   voice.name.includes("Google UK English Female") ||
//    voice.name.includes("Microsoft Zira")||
//    voice.name.includes("Samantha")
//   );
//   console.log("Set voice********", femaleVoice);
//   const speech = new SpeechSynthesisUtterance(text);
//   // speech.lang = 'en-US'; // Set language
//   // speech.rate = 1; // Speed of speech (0.1 to 10)
//   // speech.pitch = 1; // Pitch (0 to 2)
//   // speech.volume = 1; // Volume (0 to 1)
//   if (femaleVoice) {
//     speech.voice = femaleVoice;
// } else {
//     console.warn("Female voice not found. Using default voice.");
// }
//   window.speechSynthesis.speak(speech);
// }


// // Load voices properly before calling the function
// function ensureVoicesLoaded(callback) {
//   let voices = window.speechSynthesis.getVoices();
  
//   if (voices.length > 0) {
//       callback();
//   } else {
//       window.speechSynthesis.onvoiceschanged = callback;
//   }
// }


// document.getElementById('startSpeechButton').addEventListener('click', function () {
//   // Open a pop-up window to handle the speech
//   const popup = window.open('speech.html', 'SpeechRecognition', 'width=400,height=300');
// speechFlag = true;
//   // Listen for messages from the pop-up window
//   window.addEventListener("message", function (event) {
//       if (event.origin !== window.location.origin) return; // Security check

//       // Get the recognized text from the pop-up
//       const transcript = event.data;

//       // Insert recognized text into user input 
//       console.log(transcript);
//       document.getElementById("userInput").value = transcript;
//       popup.close();
//   });
// });


let femaleVoice = null; // Store the selected female voice

function loadVoices() {
    let voices = window.speechSynthesis.getVoices();
    
    // Find a female voice (adjust as needed)
    femaleVoice = voices.find(voice => 
        voice.name.includes("Samantha") ||  // macOS US Female
        voice.name.includes("Victoria") ||  // macOS AU Female
        voice.name.includes("Moira") ||     // macOS Ireland Female
        voice.name.includes("Google UK English Female") || 
        voice.name.includes("Microsoft Zira") 
    );

    console.log("Female voice selected:", femaleVoice ? femaleVoice.name : "Not found");
}

function speakText(text) {
    if (!femaleVoice) {
        console.warn("Voices not loaded yet. Using default voice.");
    }

    const speech = new SpeechSynthesisUtterance(text);
    if (femaleVoice) {
        speech.voice = femaleVoice;
    }

    window.speechSynthesis.speak(speech);
}

// Ensure voices are loaded before calling speakWithFemaleVoice
function ensureVoicesLoaded() {
    if (window.speechSynthesis.getVoices().length > 0) {
        loadVoices();
    } else {
        window.speechSynthesis.onvoiceschanged = loadVoices;
    }
}

// Load voices once when the page loads
ensureVoicesLoaded();
