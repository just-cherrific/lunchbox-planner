# Lunchbox Planner

A lightweight calendar that shows Hong Kong's current date plus the following seven days. Click a date to see its planned students and expected meals.

## Change the meal plan

Edit `public/meal-data.json`. For a known final number, set `expectedMeals`; otherwise the site estimates 65% of `students`. Commit and push the change—Vercel automatically publishes it.

## Daily refresh

`.github/workflows/daily-deploy.yml` runs at 00:05 Hong Kong time. It updates the refresh stamp and pushes a commit, which triggers Vercel to redeploy. GitHub Actions can sometimes start a few minutes late.

## Deploy in Vercel

1. Import the GitHub repository at [vercel.com/new](https://vercel.com/new).
2. Keep the Framework Preset as **Other** and leave Build Command blank.
3. Click **Deploy**. Future pushes to the connected branch deploy automatically.
