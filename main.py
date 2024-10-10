import re
import json
import time
import openai

# Function to query Perplexity API for email lookup with rate limit handling
def lookup_email(first_name, last_name, title, organization, api_key):
    # Initialize the OpenAI client for Perplexity API with the provided API key
    client = openai.OpenAI(api_key=api_key, base_url="https://api.perplexity.ai")

    messages = [
        {
            "role": "system",
            "content": (
                "You are a data analyzer. You only output in the provided JSON format. You need to help "
                "the user by providing the best possible contact information for an individual. Make sure to look deeply for the contact information!"
                "Strictly return the result in the following JSON format. Strictly don't add any other supplemental text such as comments or thoughts. Your job is to only output JSON: "
                "{'first_name': 'FirstName', 'last_name': 'LastName', 'email': 'email_address', 'phone_number': 'phone_number', 'edu_email': True/False, 'source_link': 'source_link'}"
            ),
        },
        {
            "role": "user",
            "content": (
                f"Name: {first_name} {last_name}, "
                f" Title: {title}, organization: {organization}. "
                "Prefer .edu emails, but if unavailable, provide the closest contact email. Make sure to look deeply for the contact information. "
                "Also, try to find the individual's phone number and include a link to the source of the information. "
                "Return the response as a JSON object."
            ),
        },
    ]

    # Retry mechanism with cooldown on rate limit
    while True:
        try:
            # Send the request
            response = client.chat.completions.create(
                model="llama-3.1-sonar-large-128k-online", # Uses the 70B model to process the data
                messages=messages,
            )

            # Extract the response text
            response_text = response.choices[0].message.content

            # Use regex to extract the JSON part
            json_pattern = r'\{.*\}'  # A simple pattern to capture the JSON-like object
            match = re.search(json_pattern, response_text, re.DOTALL)

            if match:
                json_string = match.group(0)  # Extract the matched JSON string
                try:
                    contact_data = json.loads(json_string)
                except json.JSONDecodeError as e:
                    print(f"Failed to decode JSON for {first_name} {last_name}: {e}")
                    contact_data = {
                        "first_name": first_name,
                        "last_name": last_name,
                        "email": None,
                        "phone_number": None,
                        "edu_email": None,
                        "source_link": None
                    }
            else:
                print(f"Failed to find JSON for {first_name} {last_name}")
                contact_data = {
                    "first_name": first_name,
                    "last_name": last_name,
                    "email": None,
                    "phone_number": None,
                    "edu_email": None,
                    "source_link": None
                }

            return contact_data

        except openai.error.RateLimitError:
            print("Rate limit exceeded for the API. Waiting for 60 seconds for the limit to reset. The program will continue afterwards.")
            time.sleep(62)  # Wait for 60 seconds before retrying
        except Exception as e:
            print(f"An error occurred: {e}")
            return {
                "first_name": first_name,
                "last_name": last_name,
                "email": None,
                "phone_number": None,
                "edu_email": None,
                "source_link": None
            }
