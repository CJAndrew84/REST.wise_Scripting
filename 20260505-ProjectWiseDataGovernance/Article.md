## Why ProjectWise Templates Don’t Scale — And What to Do About It

I was at Bentley Illuminate 2026 in Berlin and sat in a session by Benito Perez Galan on automating ProjectWise project creation.

The approach was simple:
- Define the project in Excel  
- Run a script  
- Build the ProjectWise structure  

It’s good. It works. And to be honest, most of us have been doing some version of that for years.

There were a couple of things I’ll take away from it. The way Benito structured disciplines, their folders, and the user lists tied to them was clean. I’ll use that. It makes sense.

But the core idea—Excel driving project creation—that part isn’t new.

I’ve been working like that since around 2018 when I was a Lead ProjectWise Admin for the Highways business in the UK at AECOM (what they call a GA).

Driving template setup from Excel is a clear improvement over manually building everything in ProjectWise Administrator.

Most of us already have ways to:
- Copy a template  
- Spin up a project  
- Get going quickly  

Where Benito’s approach helps is reducing the overhead. Especially around discipline folders and user lists. Define it once in the template, only create what you need. That’s solid.

I’ve worked in setups where that was already happening. The foundations were there, colleagues built a UI to request disciplines, and the deployment script only created what was needed.

That works well for getting a project up and running.

But that’s not really the problem I’m interested in anymore.

> It’s not about how we create templates.  
> It’s about how we understand them, control them, and keep track of what’s changed.

If someone joins tomorrow, can they answer:
- What does the current template actually look like?
- How did it get there?
- Why is it set up that way?

Most of the time, the answer is no.

That’s the gap.

---

[IMAGE PLACEHOLDER: Benito’s Illuminate slide showing Excel → ProjectWise automation]

---

## The Real Issue Isn’t Setup… It’s Control

ProjectWise templates in most organisations are:
- Static  
- Manually managed  
- Hard to compare  
- Hard to explain  

Which leads to:
- Inconsistent setups  
- Drift between DEV / UAT / PROD  
- No real visibility of what’s deployed  
- Governance that happens after the problem  

So yes, automate setup—but that’s only part of it.

The bigger question is:

> How do you know what you’ve actually built… and whether it’s still right?

---

Before going further, it’s worth asking:

> How many of you have inherited a ProjectWise setup you didn’t build?

Something where:
- You don’t fully understand it  
- There might be documentation  
- There probably isn’t  

That’s pretty common.

And it’s exactly where this approach helps.

---

## A Different Way to Look at It

The shift for me has been this:

> Stop treating templates as configuration.  
> Start treating them as data.

Once you do that:
- You can extract what’s actually there  
- You can compare environments  
- You can see what’s changed  
- You can start validating it  

You move from:
- “I think this is set up correctly”

to:
- “I can show exactly how this is set up”

---

## My Approach

Instead of just building templates, I extract the setup out of ProjectWise into a structured format.

At the moment, that’s Excel.

It pulls things like:
- Folder structures  
- Workflows  
- User Lists / Groups  
- Environments and attributes  
- Document codes  
- Views, rules, lookup tables  

The key point:

> This is what’s actually deployed—not what someone thinks is deployed.

---

## From Excel to Something More Useful

First step is Excel.

One workbook, multiple sheets:
- Environments  
- Workflows  
- States  
- Folders  
- Access Control  
- User Lists  
- etc  

All in one place.

That alone is useful because you can actually see it.

---

[IMAGE PLACEHOLDER: Excel workbook with multiple tabs]

---

But Excel isn’t the end game.

Next step is exporting the same thing to JSON.

Because JSON:
- Can be versioned  
- Compared  
- Queried  
- Used in pipelines  

> Excel helps you understand it.  
> JSON lets you do something with it.

---

[IMAGE PLACEHOLDER: JSON diff / version comparison]

---

## How This Works (Without Overcomplicating It)

The flow is simple:

### 1. Define it in Excel  
That’s your template definition.

### 2. Deploy to Development  
Use PowerShell to build it in a DEV datasource.

### 3. Export to JSON  
Capture what actually got deployed and store it in a repo (Azure DevOps / GitHub).

### 4. Deploy to UAT  
Let users test it properly.

### 5. Feed changes back properly  
No patching in UAT or PROD:
- Update Excel  
- Redeploy  
- Re-export JSON  
- Commit  

### 6. Deploy to Production  
Same process.

### 7. Track everything  
Git shows:
- What changed  
- When  
- Why  

It’s not about code—it’s about configuration history.

---

[IMAGE PLACEHOLDER: DEV → UAT → PROD diagram]

---

## Where This Is Going

Once the setup is structured and versioned, you can start layering things on top.

### Pipelines  
Make the process repeatable.

### Dashboards  
Actually see:
- What’s aligned  
- What’s not  
- What’s changed  

### Validation  
Check things like:
- Naming  
- Attributes  
- Workflows  
- Folder structures  
- Access  

If something’s off, flag it.

---

## Exceptions (Because There Always Are)

Not everything will follow the standard.

That’s fine.

The point isn’t to block it—it’s to make it visible.

If something doesn’t follow governance:

> You should have to explain why.

That’s it.

---

## What This Actually Changes

This isn’t about speed.

It’s about control.

- Templates → data  
- Guesswork → visibility  
- Drift → alignment  

And most importantly:

> You stop reacting to problems after they happen.

---

## Reality Check

This isn’t finished.

It depends on:
- Good data  
- Discipline  
- Knowing what “good” looks like  

Excel is just the start.

---

## Final Thought

Benito showed how to automate project creation.

That’s useful.

But the next step is:

> How do you control what you’ve built over time?

One example of where this goes:

I took one of the Excel exports and gave it to AI to review.

It picked up cascading triggers in the environments.

Something that looks fine when you’re setting it up.

But when you start pushing files through workflows, those triggers stack up and can cause performance issues.

That’s not always obvious when you’re in ProjectWise.

But when the data is structured, AI can see patterns you might miss.

---

[IMAGE PLACEHOLDER: AI response highlighting cascading triggers]

---

And this is where it gets more useful.

Once the setup is in Excel or JSON, you can ask AI to:
- Explain the template  
- Describe workflows  
- Map triggers  
- Generate flow diagrams  

You’re not digging around in ProjectWise trying to figure it out.

And more importantly:

> You’re not saying  
> “I don’t support that—I didn’t build it.”

---

## Interested to hear how others are handling this

Are you:
- Still using static templates?  
- Using Excel to drive setup?  
- Doing anything around comparison or validation?  

Genuinely interested—especially where this starts to scale.