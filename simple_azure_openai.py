from openai import AzureOpenAI
import os
import sys

import dotenv

dotenv.load_dotenv()

endpoint = os.getenv("AZURE_OPENAI_ENDPOINT", "").strip()
deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT_NAME", "gpt-5.4").strip()


def main() -> int:
    """Run a minimal Azure OpenAI chat completion smoke test."""
    api_key = os.getenv("AZURE_API_KEY", "").strip()
    if not api_key:
        print("AZURE_API_KEY is missing. Please set it in .env.")
        return 1
    if not endpoint:
        print("AZURE_OPENAI_ENDPOINT is missing. Please set it in .env.")
        return 1

    client = AzureOpenAI(
        api_version="2025-04-01-preview",
        azure_endpoint=endpoint,
        api_key=api_key,
    )

    response = client.chat.completions.create(
        messages=[
            {
                "role": "system",
                "content": "You are a helpful assistant.",
            },
            {
                "role": "user",
                "content": "What is the capital of France?",
            }
        ],
        max_completion_tokens=4096,
        # temperature=0.0,
        reasoning_effort="high",
        top_p=1.0,
        frequency_penalty=0.0,
        presence_penalty=0.0,
        model=deployment,
    )

    print(response.choices[0].message.content)
    return 0


if __name__ == "__main__":
    sys.exit(main())