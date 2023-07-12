# ReplyAI

This Outlook add-in uses Microsoft Azure OpenAI API call
At the root of the repository, create a config.js file and include your API endpoint and key
```
const env = {
  OPENAI_KEY: "[YOUR API KEY]",
  OPENAI_ENDPOINT:
    "[YOUR API ENDPOINT]",
};
export default env;
\
```
