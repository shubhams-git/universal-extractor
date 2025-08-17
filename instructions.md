I am writing a Python script, which makes a Gemini 2.5 flash call to extract the data, but Gemini doesn't accept a lot MIME type like Excel etc.
I am trying to make multiple files to handle various uploads to get converted to a pdf and later use that pdf to make a gemini call. PDF service has already been written.

Requirements: 
- Should use gemini 2.5 flash
- Should use the latest Gemini google-genai SDK 
- Should be universal in the type of files it can take (Not specific to pdf or doc or txt or any image etc) 
- Should give the result in the JSON format only. (I dont care to impose strict schema as the schema will depend upon the content of the file. But it should be able to enforce an JSON output.) 
- Should not create a virtual env If there are more information needed, please ask me before you devise your solution. 

Mandatory: Always refer to the latest documentation available for entire code that you will output. No deprecated or out-dated usage.

Make sure to refer to at least the following resource links before generating any solution.: 
https://ai.google.dev/gemini-api/docs/migrate 
https://ai.google.dev/gemini-api/docs/structured-output 
https://ai.google.dev/gemini-api/docs/text-generation 
https://ai.google.dev/gemini-api/docs/document-processing
https://ai.google.dev/gemini-api/docs/thinking 
https://ai.google.dev/gemini-api/docs/prompting-strategies