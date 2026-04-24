/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import { AppraisalData } from '../types';

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export function calculateAppraisal(data: AppraisalData) {
  // 1. Performance Appraisal Scores
  const totalEmployeeFactorScore = data.factors.reduce(
    (acc, f) => acc + (f.employeeScore * f.weightage),
    0
  );
  const totalReviewerFactorScore = data.factors.reduce(
    (acc, f) => acc + (f.reviewerScore * f.weightage),
    0
  );

  // 2. Weighted Performance Score (30% Employee, 70% Reviewer)
  const performanceScore = (totalEmployeeFactorScore * 0.3) + (totalReviewerFactorScore * 0.7);

  // 3. Goal Achievement Score
  const goalAchievementScore = data.goals.reduce(
    (acc, g) => acc + (g.score * g.weightage),
    0
  );

  // 4. Final Score (80% Performance, 20% Goals)
  const finalScore = (performanceScore * 0.8) + (goalAchievementScore * 0.2);

  return {
    employeeFactorScore: totalEmployeeFactorScore,
    reviewerFactorScore: totalReviewerFactorScore,
    performanceScore,
    goalAchievementScore,
    finalScore,
    ratingCategory: getRatingCategory(finalScore)
  };
}

export function getRatingCategory(score: number) {
  if (score === 0) return "N/A";
  if (score < 2.0) return "Poor";
  if (score < 3.0) return "Needs Improvement";
  if (score < 4.0) return "Good";
  if (score < 4.6) return "Very Good";
  return "Exceptional";
}

export function formatScore(score: number) {
  return score.toFixed(2);
}
