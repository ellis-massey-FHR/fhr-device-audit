pipeline {
  agent {
    docker {
      image 'python:3.11'
      args '-v /etc/timezone:/etc/timezone:ro -v /etc/localtime:/etc/localtime:ro'
      reuseNode true
    }
  }

  parameters {
    string(name: 'SCRIPT_PATH', defaultValue: 'scripts/get_servicenow_data.py', description: 'Path to the Python script in the repo')
    string(name: 'SCRIPT_ARGS', defaultValue: '', description: 'Optional args to pass to the script')
  }

  environment { TZ = 'America/Chicago' }
  triggers { cron('H 7 * * 1-5') }  // runs on weekdays; you can remove/adjust

  options { timestamps(); disableConcurrentBuilds(); buildDiscarder(logRotator(numToKeepStr: '20')) }

  stages {
    stage('Checkout'){ steps { checkout scm } }

    stage('Setup Python'){
      steps {
        sh '''
          python --version
          python -m pip install --upgrade pip
          [ -f requirements.txt ] && pip install -r requirements.txt || true
        '''
      }
    }

    stage('Run'){
      steps {
        sh '''
          set -e
          mkdir -p output
          if [ ! -f "$SCRIPT_PATH" ]; then
            echo "❌ SCRIPT_PATH not found: $SCRIPT_PATH"
            echo "Repo layout:"
            ls -R | sed -n '1,200p'
            exit 1
          fi
          echo "▶ Running: python $SCRIPT_PATH $SCRIPT_ARGS"
          python "$SCRIPT_PATH" $SCRIPT_ARGS
        '''
      }
    }

    stage('Archive'){
      when { expression { fileExists('output') } }
      steps { archiveArtifacts artifacts: 'output/**', fingerprint: true, allowEmptyArchive: true }
    }
  }

  post {
    success { echo '✅ Job done.' }
    failure { echo '❌ Build failed. Check Console Output.' }
  }
}
