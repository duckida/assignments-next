/** Demo dataset generated relative to "now" so the UI always looks alive.
 * Mirrors the raw-ish Graph response shape (assignment + submissions map) that
 * the service worker returns, so demo and live paths share the whole pipeline.
 */

const H = 3_600_000
const D = 86_400_000

export function buildDemoData(now = Date.now()) {
  const classes = [
    { id: 'cls-bio-3d1f', name: 'Biology 12 · P3', description: 'SCI-12-03', webUrl: 'https://teams.microsoft.com/' },
    { id: 'cls-calc-8a02', name: 'AP Calculus AB', description: 'MTH-AP-CA', webUrl: 'https://teams.microsoft.com/' },
    { id: 'cs-lit-525', name: 'English Literature 11', description: 'ENG-LIT-11', webUrl: 'https://teams.microsoft.com/' },
    { id: 'cls-cs-77b1', name: 'Computer Science 20', description: 'CS-20-01', webUrl: 'https://teams.microsoft.com/' },
    { id: 'cls-hist-c441', name: 'Modern World History', description: 'HIS-MW-10', webUrl: 'https://teams.microsoft.com/' },
  ]

  const assignments = [
    {
      id: 'as-001', classId: 'cls-bio-3d1f', displayName: 'Lab Report: Osmosis & Diffusion',
      dueDateTime: new Date(now - 2 * D - 5 * H).toISOString(), assignedDateTime: new Date(now - 9 * D).toISOString(),
      status: 'assigned', allowLateSubmissions: true, grading: { '@odata.type': '#microsoft.graph.educationAssignmentPointsGradeType', maxPoints: 25 },
      webUrl: 'https://teams.microsoft.com/', lastModifiedDateTime: new Date(now - 2 * D).toISOString(),
    },
    {
      id: 'as-002', classId: 'cls-calc-8a02', displayName: 'Problem Set 9 — Related Rates',
      dueDateTime: new Date(now + 3 * H).toISOString(), assignedDateTime: new Date(now - 6 * D).toISOString(),
      status: 'assigned', allowLateSubmissions: false, grading: { '@odata.type': '#microsoft.graph.educationAssignmentPointsGradeType', maxPoints: 20 },
      webUrl: 'https://teams.microsoft.com/', lastModifiedDateTime: new Date(now - D).toISOString(),
    },
    {
      id: 'as-003', classId: 'cs-lit-525', displayName: 'Task: Chapter 10 reflection — "The Hound"',
      dueDateTime: new Date(now + 26 * H).toISOString(), assignedDateTime: new Date(now - 3 * D).toISOString(),
      status: 'assigned', allowLateSubmissions: true, grading: { '@odata.type': '#microsoft.graph.educationAssignmentCompletionGradeType' },
      webUrl: 'https://teams.microsoft.com/', lastModifiedDateTime: new Date(now - 3 * D).toISOString(),
    },
    {
      id: 'as-004', classId: 'cls-cs-77b1', displayName: 'Unit 4 Checkpoint: Recursion',
      dueDateTime: new Date(now + 3 * D + 4 * H).toISOString(), assignedDateTime: new Date(now - D).toISOString(),
      status: 'assigned', allowLateSubmissions: true, grading: { '@odata.type': '#microsoft.graph.educationAssignmentPointsGradeType', maxPoints: 30 },
      webUrl: 'https://teams.microsoft.com/', lastModifiedDateTime: new Date(now - 11 * H).toISOString(),
    },
    {
      id: 'as-005', classId: 'cls-hist-c441', displayName: 'Source Analysis: Treaty of Versailles (1919)',
      dueDateTime: new Date(now + 5 * D).toISOString(), assignedDateTime: new Date(now - 4 * D).toISOString(),
      status: 'assigned', allowLateSubmissions: false, grading: { '@odata.type': '#microsoft.graph.educationAssignmentPointsGradeType', maxPoints: 15 },
      webUrl: 'https://teams.microsoft.com/', lastModifiedDateTime: new Date(now - 29 * H).toISOString(),
    },
    {
      id: 'as-006', classId: 'cls-calc-8a02', displayName: 'Formative: Chain rule drill',
      dueDateTime: new Date(now + 5 * D + 2 * H).toISOString(), assignedDateTime: new Date(now - 2 * D).toISOString(),
      status: 'assigned', allowLateSubmissions: true, grading: { '@odata.type': '#microsoft.graph.educationAssignmentCompletionGradeType' },
      webUrl: 'https://teams.microsoft.com/', lastModifiedDateTime: new Date(now - 47 * H).toISOString(),
    },
    {
      id: 'as-007', classId: 'cls-bio-3d1f', displayName: 'Pre-lab: Kidney dissection hygiene sheet',
      dueDateTime: new Date(now + 44 * H).toISOString(), assignedDateTime: new Date(now - 1 * D).toISOString(),
      status: 'assigned', allowLateSubmissions: false, grading: { '@odata.type': '#microsoft.graph.educationAssignmentCompletionGradeType' },
      webUrl: 'https://teams.microsoft.com/', lastModifiedDateTime: new Date(now - 20 * H).toISOString(),
    },
    {
      id: 'as-008', classId: 'cls-cs-77b1', displayName: 'Design doc — Portfolio site IA',
      dueDateTime: new Date(now + 9 * D).toISOString(), assignedDateTime: new Date(now - 5 * D).toISOString(),
      status: 'assigned', allowLateSubmissions: true, grading: { '@odata.type': '#microsoft.graph.educationAssignmentPointsGradeType', maxPoints: 10 },
      webUrl: 'https://teams.microsoft.com/', lastModifiedDateTime: new Date(now - 5 * D).toISOString(),
    },
    {
      id: 'as-009', classId: 'cs-lit-525', displayName: 'Essay draft: identity in "Frankenstein"',
      dueDateTime: new Date(now + 12 * D).toISOString(), assignedDateTime: new Date(now - 5 * D).toISOString(),
      status: 'assigned', allowLateSubmissions: true, grading: { '@odata.type': '#microsoft.graph.educationAssignmentPointsGradeType', maxPoints: 50 },
      webUrl: 'https://teams.microsoft.com/', lastModifiedDateTime: new Date(now - 5 * D).toISOString(),
    },
    {
      id: 'as-010', classId: 'cls-calc-8a02', displayName: 'Problem Set 8 — Implicit differentiation',
      dueDateTime: new Date(now - 4 * D).toISOString(), assignedDateTime: new Date(now - 11 * D).toISOString(),
      status: 'assigned', allowLateSubmissions: true, grading: { '@odata.type': '#microsoft.graph.educationAssignmentPointsGradeType', maxPoints: 20 },
      webUrl: 'https://teams.microsoft.com/', lastModifiedDateTime: new Date(now - 4 * D).toISOString(),
    },
    {
      id: 'as-011', classId: 'cls-hist-c441', displayName: 'Timeline cards: interwar period',
      dueDateTime: new Date(now - 8 * D).toISOString(), assignedDateTime: new Date(now - 15 * D).toISOString(),
      status: 'assigned', allowLateSubmissions: true, grading: { '@odata.type': '#microsoft.graph.educationAssignmentCompletionGradeType' },
      webUrl: 'https://teams.microsoft.com/', lastModifiedDateTime: new Date(now - 8 * D).toISOString(),
    },
    {
      id: 'as-012', classId: 'cls-bio-3d1f', displayName: 'Quiz response: cell cycle',
      dueDateTime: new Date(now - 7 * D).toISOString(), assignedDateTime: new Date(now - 14 * D).toISOString(),
      status: 'assigned', allowLateSubmissions: false, grading: { '@odata.type': '#microsoft.graph.educationAssignmentPointsGradeType', maxPoints: 15 },
      webUrl: 'https://teams.microsoft.com/', lastModifiedDateTime: new Date(now - 7 * D).toISOString(),
    },
    {
      id: 'as-013', classId: 'cls-cs-77b1', displayName: 'Worksheet: binary trees',
      dueDateTime: new Date(now - 6 * D).toISOString(), assignedDateTime: new Date(now - 10 * D).toISOString(),
      status: 'assigned', allowLateSubmissions: true, grading: { '@odata.type': '#microsoft.graph.educationAssignmentPointsGradeType', maxPoints: 12 },
      webUrl: 'https://teams.microsoft.com/', lastModifiedDateTime: new Date(now - 6 * D).toISOString(),
    },
  ]

  const submissions = {
    'as-001': { id: 'sub-001', status: 'working', resourceIds: [] },
    'as-002': { id: 'sub-002', status: 'working', resourceIds: [] },
    'as-003': { id: 'sub-003', status: 'working', resourceIds: [] },
    'as-004': { id: 'sub-004', status: 'working', resourceIds: [] },
    'as-005': { id: 'sub-005', status: 'working', resourceIds: [] },
    'as-006': { id: 'sub-006', status: 'working', resourceIds: [] },
    'as-007': { id: 'sub-007', status: 'working', resourceIds: [] },
    'as-008': { id: 'sub-008', status: 'working', resourceIds: [] },
    'as-009': { id: 'sub-009', status: 'working', resourceIds: [] },
    'as-010': {
      id: 'sub-010', status: 'returned',
      submittedDateTime: new Date(now - 5 * D).toISOString(), returnedDateTime: new Date(now - 3 * D).toISOString(),
      outcomes: [
        { '@odata.type': '#microsoft.graph.educationPointsOutcome', points: { '@odata.type': '#microsoft.graph.educationPointsOutcomePoints', points: 17, publishedPoints: 17 } },
        { '@odata.type': '#microsoft.graph.educationFeedbackOutcome', feedback: { '@odata.type': '#microsoft.graph.educationFeedback', text: '<p>Nice handling of the implicit cases — check the sign on <strong>&minus;6x&sup3;</strong> in step 3 and make sure your final line includes units.</p>' } },
      ],
    },
    'as-011': {
      id: 'sub-011', status: 'submitted',
      submittedDateTime: new Date(now - 9 * D).toISOString(),
      outcomes: [],
    },
    'as-012': {
      id: 'sub-012', status: 'returned',
      submittedDateTime: new Date(now - 8 * D).toISOString(), returnedDateTime: new Date(now - 6 * D).toISOString(),
      outcomes: [
        { '@odata.type': '#microsoft.graph.educationPointsOutcome', points: { '@odata.type': '#microsoft.graph.educationPointsOutcomePoints', points: 14, publishedPoints: 14 } },
        { '@odata.type': '#microsoft.graph.educationFeedbackOutcome', feedback: { '@odata.type': '#microsoft.graph.educationFeedback', text: '<p>Solid recall of the checkpoints. Revisit the metaphase/anaphase distinction for the unit test.</p>' } },
      ],
    },
    'as-013': {
      id: 'sub-013', status: 'reassigned',
      submittedDateTime: new Date(now - 7 * D).toISOString(), reassignedDateTime: new Date(now - 2 * D).toISOString(),
      outcomes: [
        { '@odata.type': '#microsoft.graph.educationFeedbackOutcome', feedback: { '@odata.type': '#microsoft.graph.educationFeedback', text: '<p>The traversal is close — redo Q3 with your corrected swap and resubmit.</p>' } },
      ],
    },
  }

  const details = {
    // everything falls back to a generic detail sheet; a few have instructions + resources
    'as-004': {
      instructions: { content: '<div class="grid-directions"><ol><li>Complete <strong>Q1–Q4</strong> on paper.</li><li>Submit a single PDF scan plus your <code>.cs</code> file.</li><li>Late penalty: 5% per day after 4:00 PM.</li></ol></div>' },
      resources: [
        { id: 'rsp-1', displayName: 'checkpoint-spec.pdf', assignmentResourceType: 'fileUrl', resource: { '@odata.type': '#microsoft.graph.educationLinkResource', link: 'https://example.com/spec.pdf' } },
        { id: 'rsp-2', displayName: 'Starter code (GitHub Classroom)', assignmentResourceType: 'link', resource: { '@odata.type': '#microsoft.graph.educationLinkResource', link: 'https://github.com/' } },
      ],
    },
  }

  return { classes, assignments, submissions, details }
}

/** adapter: build the same records the SW returns from demo data */
export function demoDetail(id) {
  const d = buildDemoData()
  const a = d.assignments.find((x) => x.id === id)
  if (!a) return null
  return {
    ...a,
    instructions: d.details[id]?.instructions || { content: '' },
    resources: d.details[id]?.resources || [],
  }
}
