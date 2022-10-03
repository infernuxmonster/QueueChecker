# QueueChecker

QueueChecker is a small project created in PowerShell to allow World of Warcraft players to monitor the queue on their realm.

Similar to my other project, [WorldBoss Announcer](https://github.com/infernuxmonster/Worldboss-Announcer) it uses PowerShell, OCR and Discord to set up monitoring for queues in World of Warcraft Classic.

## Prereq

1. [OCR Space API Key](https://ocr.space/ocrapi/freekey)
2. [Your own Discord server webhook](https://support.discord.com/hc/en-us/articles/228383668-Intro-to-Webhooks)
3. Powershell

## Running the script

After getting the OCR Space Api Key and a Discord webhook, you should be able to run the script by right-clicking and selecting "Run with Powershell".

![](QueueChecker.png)

## Common errors

You might get an error about the execution policy - I do recommend only running scripts you know what do, and in that case you can run `Set-ExecutionPolicy Bypass` to allow the script to run.
