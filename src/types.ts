/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface EmployeeInfo {
    name: string;
    position: string;
    department: string;
    projectsManaged: string;
    appraisalType: string;
    appraisalPeriod: string;
  }
  
  export interface ReviewerInfo {
    name: string;
    position: string;
    department: string;
    appraisalDue: string;
  }
  
  export interface PerformanceFactor {
    id: string;
    title: string;
    description: string;
    weightage: number;
    employeeScore: number;
    reviewerScore: number;
    comments: string;
  }
  
  export interface Goal {
    id: string;
    description: string;
    score: number;
    weightage: number;
    comments: string;
  }
  
  export interface AppraisalData {
    employee: EmployeeInfo;
    reviewer: ReviewerInfo;
    factors: PerformanceFactor[];
    goals: Goal[];
    comments: {
      employee: string;
      reviewer: string;
      ceo: string;
    };
    signatures: {
      employee: { name: string; date: string };
      reviewer: { name: string; date: string };
      ceo: { name: string; date: string };
    };
  }
  
  export const INITIAL_FACTORS: PerformanceFactor[] = [
    {
      id: "01",
      title: "Subject Knowledge / Expertise",
      description: "Consider the employee's skill level, knowledge and understanding of all phases of the job and those requiring improved skills and/or experience.",
      weightage: 0.15,
      employeeScore: 0,
      reviewerScore: 0,
      comments: "",
    },
    {
      id: "02",
      title: "Work Execution & Coordination",
      description: "Measures effectiveness in planning, organizing, and efficiently handling tasks and activities assigned on time, while eliminating unnecessary activities.",
      weightage: 0.15,
      employeeScore: 0,
      reviewerScore: 0,
      comments: "",
    },
    {
      id: "03",
      title: "Communication (Internal & Client)",
      description: "Effective in listening, expressing ideas, and providing relevant and timely information to management, coworkers, subordinates and customers.",
      weightage: 0.15,
      employeeScore: 0,
      reviewerScore: 0,
      comments: "",
    },
    {
      id: "04",
      title: "Problem Solving",
      description: "Measures effectiveness in understanding problems and making suggestions for rectifying the issue, and apply timely, practical decisions.",
      weightage: 0.15,
      employeeScore: 0,
      reviewerScore: 0,
      comments: "",
    },
    {
      id: "05",
      title: "Teamwork",
      description: "Measures how well this individual gets along with fellow employees and able to work cooperatively to get work done.",
      weightage: 0.15,
      employeeScore: 0,
      reviewerScore: 0,
      comments: "",
    },
    {
      id: "06",
      title: "Independent Action",
      description: "Measures effectiveness in time management; initiative and independent action within prescribed limits.",
      weightage: 0.05,
      employeeScore: 0,
      reviewerScore: 0,
      comments: "",
    },
    {
      id: "07",
      title: "Dependability",
      description: "How well employees comply with instructions and perform under unusual circumstances; consider record of attendance and punctuality.",
      weightage: 0.05,
      employeeScore: 0,
      reviewerScore: 0,
      comments: "",
    },
    {
      id: "08",
      title: "Learning Initiatives",
      description: "Consider the initiatives taken by the employee in pursuing upskilling opportunities within own limits (i.e., online resources, courses).",
      weightage: 0.05,
      employeeScore: 0,
      reviewerScore: 0,
      comments: "",
    },
    {
      id: "09",
      title: "Innovation",
      description: "Contribution to new ideas, process improvements suggested, and innovation mindset. Proactively seeks better ways to achieve goals.",
      weightage: 0.10,
      employeeScore: 0,
      reviewerScore: 0,
      comments: "",
    },
  ];
  
  export const INITIAL_GOALS: Goal[] = [
    { id: "01", description: "", score: 0, weightage: 0.6, comments: "" },
    { id: "02", description: "", score: 0, weightage: 0.2, comments: "" },
    { id: "03", description: "", score: 0, weightage: 0.2, comments: "" },
  ];
  