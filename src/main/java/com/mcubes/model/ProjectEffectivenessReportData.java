package com.mcubes.model;

public class ProjectEffectivenessReportData {

    private String workspaceName;
    private String projectName;
    private long populationSize;
    private int recallTarget;
    private int probabilityTargetAchieved;
    private long reviewedElusionSampleCreation;
    private long foundElusionSampleCreation;
    private long reviewedStoppingPoint;
    private long foundStoppingPoint;
    private int elusionSampleSize;
    private int relevantDocsInSample;


    public String getWorkspaceName() {
        return workspaceName;
    }

    public void setWorkspaceName(String workspaceName) {
        this.workspaceName = workspaceName;
    }

    public String getProjectName() {
        return projectName;
    }

    public void setProjectName(String projectName) {
        this.projectName = projectName;
    }

    public long getPopulationSize() {
        return populationSize;
    }

    public void setPopulationSize(long populationSize) {
        this.populationSize = populationSize;
    }

    public int getRecallTarget() {
        return recallTarget;
    }

    public void setRecallTarget(int recallTarget) {
        this.recallTarget = recallTarget;
    }

    public int getProbabilityTargetAchieved() {
        return probabilityTargetAchieved;
    }

    public void setProbabilityTargetAchieved(int probabilityTargetAchieved) {
        this.probabilityTargetAchieved = probabilityTargetAchieved;
    }

    public long getReviewedElusionSampleCreation() {
        return reviewedElusionSampleCreation;
    }

    public void setReviewedElusionSampleCreation(long reviewedElusionSampleCreation) {
        this.reviewedElusionSampleCreation = reviewedElusionSampleCreation;
    }

    public long getFoundElusionSampleCreation() {
        return foundElusionSampleCreation;
    }

    public void setFoundElusionSampleCreation(long foundElusionSampleCreation) {
        this.foundElusionSampleCreation = foundElusionSampleCreation;
    }

    public long getReviewedStoppingPoint() {
        return reviewedStoppingPoint;
    }

    public void setReviewedStoppingPoint(long reviewedStoppingPoint) {
        this.reviewedStoppingPoint = reviewedStoppingPoint;
    }

    public long getFoundStoppingPoint() {
        return foundStoppingPoint;
    }

    public void setFoundStoppingPoint(long foundStoppingPoint) {
        this.foundStoppingPoint = foundStoppingPoint;
    }

    public int getElusionSampleSize() {
        return elusionSampleSize;
    }

    public void setElusionSampleSize(int elusionSampleSize) {
        this.elusionSampleSize = elusionSampleSize;
    }

    public int getRelevantDocsInSample() {
        return relevantDocsInSample;
    }

    public void setRelevantDocsInSample(int docsInElusionSample) {
        this.relevantDocsInSample = docsInElusionSample;
    }
}
