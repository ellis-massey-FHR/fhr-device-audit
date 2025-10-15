pipeline {
  agent {
    docker {
      image 'python:3.11'
      args '-v /etc/timezone:/etc/timezone:ro -v /etc/localtime:/etc/localtime:ro'
      reuseNode true
    }
  }

  environment {
    TZ = 'America/Chicago'
    // Example for secrets later:
    // API_TOKEN = credentials('my-api-token-id')
  }

  // Run on weekdays near 7 AM CT; remove or edit as you like
  triggers { cron('H 7 * * 1-5') }

  options {
    timestamps()
    disableConcurrentBuilds()
    buildDiscarder(logRotator(numToKeepStr: '20'))
  }

  stages {
    stage('Checkout') { steps { checkout scm } }

    stage('Setup Python') {
      steps {
        sh '''
          python --version
          python -m pip install --upgrade pip
          if [ -f requirements.txt ]; then pip install -r requirements.txt; fi
        '''
      }
    }

    stage('Run Task') {
      steps {
        sh '''
          set -e
          mkdir -p output
          if [ -f scripts/my_task.py ]; then
            python scripts/my_task.py
          elif [ -f main.py ]; then
            python main.py
          else
            echo "No known entrypoint. Update Jenkinsfile."; exit 1
          fi
        '''
      }
    }

    stage('Archive Outputs') {
      when { expression { fileExists('output') } }
      steps { archiveArtifacts artifacts: 'output/**', fingerprint: true }
    }
  }

  post {
    success { echo '✅ Job done.' }
    failure { echo '❌ Build failed. Check Console Output.' }
  }
}
