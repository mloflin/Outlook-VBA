# Outlook-VBA

The focus of this automation is to automatically schedule (1) Lunch, (2) Time to Cover Email, and (3) Be able to quickly take an email and schedule a block of time to review it.

What I found missing from setting up regularly recurring meetings was that they didn't allow for flexibility and I spent more time shuffling them around than regularly having lunch/reviewing email. 

I have this script running everytime a meeting reminder is triggered to look out 7 days from now and auto-schedule the invites based on some logic (below). The result is that I have varying lunch schedules and random amounts of time to review email. Freeing me up to have more time working on projects.

- Lunch: It looks 7 days out and automatically schedules a 15, 30, or 60 minute window from between 11:30am-1pm based on availability
- Email: It looks 7 days out and automatically schedules a 15, 30, or 60 minute window for Email once a day either from 9-11:30 or from 1-4pm based on availability

It stores some local variables to be able to accomodate for skipped days and will automatically schedule any missed days. It skips weekends and for Lunch, I have it adding my personal shared calendar. Once it runs for a given day, it doesn't schedule again until the next day. It also has a random concept for Email where it randomly selects 60, 30, or 15 minutes instead of always being 60 minutes. If there are no 60 then it tries 30 and then tries 15.

For #3 - I have a couple of Macro Buttons in my top bar that allow me to select a given email and then push the given Macro Button and it automatically Archives the email and schedules an available slot of time (at least <X> hours in the future) to review the content of said email. 
- Example would be: Schedule a block of 15 minutues to review <Y> email

Known issues:
- It sometimes slows down reminders given it is running the code in the background.
