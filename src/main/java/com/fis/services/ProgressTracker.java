package com.fis.services;

public class ProgressTracker {
    private int totalTasks;
    private int completedTasks;

    public ProgressTracker(int totalTasks) {
        this.totalTasks = totalTasks;
        this.completedTasks = 0;
    }

    public synchronized void taskCompleted() {
        completedTasks++;
    }

    public synchronized int getProgressPercentage() {
        return (int) ((completedTasks / (double) totalTasks) * 100);
    }

    public synchronized String getProgressBar() {
        int progress = (int) ((completedTasks / (double) totalTasks) * 100);
        StringBuilder progressBar = new StringBuilder("[");
        for (int i = 0; i < progress / 10; i++) {
            progressBar.append("=");
        }
        for (int i = progress / 10; i < 10; i++) {
            progressBar.append(" ");
        }
        progressBar.append("]");
        return progressBar.toString();
    }
}
